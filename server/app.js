const express = require("express"), fs = require('fs'), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => { const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`]; allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization')); next(); });
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => { console.log(`Server running at http://${hostname}:${port}/`); });

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => { lastUploadedFilePath = p; const workbook = xlsx.readFile(p), sheet = workbook.Sheets[workbook.SheetNames[0]], json = xlsx.utils.sheet_to_json(sheet); res.json({ agentNames: [...new Set(json.map(data => data["Taken By"]))] }); });
app.post("/process", upload.single("file"), async ({ body: config }, res) => { try { const workbook = xlsx.readFile(lastUploadedFilePath), sheet = workbook.Sheets[workbook.SheetNames[0]], originalXlData = xlsx.utils.sheet_to_json(sheet), { incidentConfigs, sfMembers, incidentsPerAgent } = config, selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, incidentsPerAgent, mapSFMembersToIncidentAgents(sfMembers, mapIncidentsByAgent(originalXlData))), rows = formatRowsForDownload(selectedIncidents); if (rows.length < incidentsPerAgent) { throw new Error("Not enough incidents matched the provided configuration"); } else { downloadFile(res, createAndWriteWorksheet(workbook, rows)); }; console.log(config); } catch (error) { console.error("Error in /process:", error, "Request body:", config, lastUploadedFilePath && "Last uploaded file path:", lastUploadedFilePath); res.status(500).send("Internal Server Error"); } });

const {
  getRandomValue = (incidents, field, alreadySelected) => { const uniqueValues = [...new Set(incidents.map((i) => i[field]))]; const unselectedValues = uniqueValues.filter((value) => !alreadySelected.has(value)); const valuesToUse = unselectedValues.length > 0 ? unselectedValues : uniqueValues; return valuesToUse[Math.floor(Math.random() * valuesToUse.length)]; },
  fisherYatesShuffle = (array) => { array.forEach((_, i) => { const j = Math.floor(Math.random() * (i + 1));[array[i], array[j]] = [array[j], array[i]]; }); return array; },
  mapIncidentsByAgent = (data) => data.reduce((acc, incident) => { const agent = incident["Taken By"]; acc[agent] = acc[agent] || []; acc[agent].push(incident); return acc; }, {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => { const sfAgentMapping = {}, agents = Object.keys(incidentsByAgent); fisherYatesShuffle(agents).forEach((agent, index) => { const sfMember = sfMembers[index % sfMembers.length]; sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent]; }); return sfAgentMapping; },
  selectUniqueIncident = (filteredIncidents, alreadySelected) => { const uniqueIncidents = filteredIncidents.filter((incident) => !alreadySelected.has(incident)); return uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : null; },
  selectUniqueIncidentForAgent = (filteredIncidents = [], alreadySelected, originalIncidents = []) => { let uniqueIncidents = filteredIncidents.filter((incident) => !alreadySelected.has(incident)); let selectedIncident; selectedIncident = uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : (() => { let remainingOriginals = originalIncidents.filter((incident) => !alreadySelected.has(incident)); return remainingOriginals.length ? remainingOriginals[Math.floor(Math.random() * remainingOriginals.length)] : (console.warn('All original incidents have been selected for the current agent.'), null); })(); selectedIncident ? alreadySelected.add(selectedIncident) : null; return selectedIncident; },
  filterIncidentsByCriterion = (incidents, field, value, agent, alreadySelected, originalIncidents) => { const filtered = filterByCriterion(incidents, field, value, agent, alreadySelected, originalIncidents); const selectedIncident = selectUniqueIncident(filtered, alreadySelected); if (!selectedIncident && value !== 'RANDOM') { const fallbackFiltered = originalIncidents.filter(incident => incident[field] === value && !alreadySelected.has(incident)); return fallbackFiltered.length ? [fallbackFiltered[Math.floor(Math.random() * fallbackFiltered.length)]] : []; } return selectedIncident ? [selectedIncident] : []; },
  filterByCriterion = (incidents, field, value, agent, alreadySelected, originalIncidents) => { value = value === "RANDOM" ? getRandomValue(incidents, field, alreadySelected) : value; const filtered = incidents.filter(incident => { const matches = !alreadySelected.has(incident) && incident[field] === value && incident["Taken By"] === agent; if (!matches && incident[field] === value && incident["Taken By"] === agent) { console.log(`Excluded incident: ${incident}, field: ${field}, value: ${value}, agent: ${agent}`); } return matches; }); if (!filtered.length) { switch (field) { case 'Service': return [selectUniqueIncidentForAgent(originalIncidents, alreadySelected)]; case 'Contact type': const contactTypeFiltered = originalIncidents.filter(incident => incident['Contact type'] === value); return [selectUniqueIncidentForAgent(contactTypeFiltered, alreadySelected, originalIncidents)]; case 'First time fix': return [selectUniqueIncidentForAgent(incidents, alreadySelected)]; default: console.log(`Warning: No incidents available for fallback for agent ${agent}.`); return []; } } return filtered; },
  selectIncidentsByConfiguration = async (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) => { (!Array.isArray(originalXlData) || !originalXlData.length) && (() => { throw new Error("Invalid originalXlData"); })(); return Object.entries(sfAgentMapping).reduce((selectedIncidents, [sfMember, agents]) => { selectedIncidents[sfMember] = agents.reduce((agentIncidents, agent) => { const alreadySelected = new Set(); agentIncidents[agent] = Array(maxIncidents).fill().reduce((incidents, _, i) => { const incidentConfig = incidentConfigs[i % incidentConfigs.length]; let potentialIncidents = [...originalXlData];["Service", "Contact type", "First time fix"].forEach(field => { const filtered = incidentConfig[field.toLowerCase()] ? filterIncidentsByCriterion(potentialIncidents, field, incidentConfig[field.toLowerCase()], agent, alreadySelected, originalXlData) : []; potentialIncidents = filtered.length ? filtered : potentialIncidents; if (!filtered.length && incidents.length <= i) { const randomIncident = selectUniqueIncidentForAgent(potentialIncidents, alreadySelected); incidents.push(randomIncident); alreadySelected.add(randomIncident.incidentId); } }); incidents.length <= i && potentialIncidents.length ? ((uniqueIncident = selectUniqueIncidentForAgent(potentialIncidents, alreadySelected)), incidents.push(uniqueIncident), alreadySelected.add(uniqueIncident.incidentId)) : console.log(`Warning: No incidents available for fallback for agent ${agent}.`); return incidents; }, []); return agentIncidents; }, {}); return selectedIncidents; }, {}); },
} = {};

/* WRITE + DOWNLOAD */
const { createAndWriteWorksheet = (workbook, rows) => { const newWorksheet = xlsx.utils.json_to_sheet(rows); workbook.Sheets["Processed List"] ? (workbook.Sheets["Processed List"] = newWorksheet) : xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List"); const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME); xlsx.writeFile(workbook, newFilePath); return newFilePath; }, formatRowsForDownload = (selectedIncidents) => { let [previousSFMember, previousAgent] = [""]; return Object.keys(selectedIncidents).flatMap((sfMember) => Object.keys(selectedIncidents[sfMember]).flatMap((agent) => selectedIncidents[sfMember][agent].map((incident) => { const row = { "SF Member": previousSFMember === sfMember ? "" : sfMember, Agent: previousAgent === agent ? "" : agent, "Task Number": incident["Task Number"], Service: incident["Service"], "Contact type": incident["Contact type"], "First time fix": incident["First time fix"], };[previousSFMember, previousAgent] = [sfMember, agent]; return row; }))); }, downloadFile = (res, newFilePath) => { res.download(newFilePath, filename, (err) => { if (err) throw new Error("Error sending the file: " + err); fs.unlink(newFilePath, (err) => { if (err) throw new Error("Error deleting the processed file: " + err); console.log("Processed file deleted successfully"); }); fs.unlink(lastUploadedFilePath, (err) => { if (err) throw new Error("Error deleting the temporary file: " + err); console.log("Temporary file deleted successfully"); }); }); } } = {};