const express = require("express"), cors = require('cors'), fs = require('fs'), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(cors());
app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => { const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`]; allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization')); next(); });
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => { console.log(`Server running at http://${hostname}:${port}/`); });

/* ENDPOINTS */

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => { lastUploadedFilePath = p; const workbook = xlsx.readFile(p), sheet = workbook.Sheets[workbook.SheetNames[0]], json = xlsx.utils.sheet_to_json(sheet); res.json({ agentNames: [...new Set(json.map(data => data["Taken By"]))] }); });
app.post("/process", upload.single("file"), async ({ body: config }, res) => { try { const workbook = xlsx.readFile(lastUploadedFilePath), sheet = workbook.Sheets[workbook.SheetNames[0]], originalXlData = xlsx.utils.sheet_to_json(sheet), { incidentConfigs, sfMembers, incidentsPerAgent } = config, selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, incidentsPerAgent, mapSFMembersToIncidentAgents(sfMembers, mapIncidentsByAgent(originalXlData))), rows = formatRowsForDownload(selectedIncidents); if (rows.length < incidentsPerAgent) { throw new Error("Not enough incidents matched the provided configuration"); } else { downloadFile(res, createAndWriteWorksheet(workbook, rows)); }; } catch (error) { console.error("Request body:", config, "Last uploaded file path:", lastUploadedFilePath, "Error in /process:", error); res.status(500).send("Internal Server Error"); } });

/* PROCESS DATA */

const {

  getRandomValue = (incidents, field, alreadySelected) => (values => values[Math.floor(Math.random() * values.length)])(([...new Set(incidents.map(i => i[field]))].filter(value => !alreadySelected.has(value))) || [...new Set(incidents.map(i => i[field]))]),
  fisherYatesShuffle = array => array.reduceRight((_, __, i, arr) => ((j) => ([arr[i], arr[j]] = [arr[j], arr[i]], arr))(Math.floor(Math.random() * (i + 1)))),
  mapIncidentsByAgent = data => data.reduce((acc, { ["Taken By"]: agent, ...incident }) => ({ ...acc, [agent]: [...(acc[agent] || []), incident] }), {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => fisherYatesShuffle(Object.keys(incidentsByAgent)).reduce((sfAgentMapping, agent, index) => ({ ...sfAgentMapping, [sfMembers[index % sfMembers.length]]: [...(sfAgentMapping[sfMembers[index % sfMembers.length]] || []), agent] }), {}),
  selectUniqueIncidentForAgent = (filteredIncidents = [], alreadySelected, originalIncidents = []) => { const [getUniqueIncidents, getRandomIncident] = [incidents => incidents.filter(incident => !alreadySelected.has(incident["Task Number"])), incidents => incidents[~~(Math.random() * incidents.length)]]; const selectedIncident = getRandomIncident((filteredIncidents.length ? filteredIncidents : originalIncidents).filter(incident => !alreadySelected.has(incident["Task Number"]))); selectedIncident && (console.warn(selectedIncident ? "" : "All original incidents have been selected for the current agent."), alreadySelected.add(selectedIncident["Task Number"])); return selectedIncident; },

  filterByCriterion = (incidents, field, value, agent, alreadySelected, triedValues = new Set()) => {
    const helper = (value, triedValues) => {
      triedValues.add(value);
      const filtered = incidents.filter((incident) => !alreadySelected.has(incident) && String(incident[field]).toLowerCase() === (typeof value === 'string' ? value.toLowerCase() : value) && incident["Taken By"] === agent);
      const untriedValues = incidents.map((i) => i[field]).filter((v) => !triedValues.has(v));
      if (field === 'First time fix' && filtered.length === 0 && !triedValues.has(value === "TRUE" ? "FALSE" : "TRUE")) {
        return helper(value === "TRUE" ? "FALSE" : "TRUE", triedValues);
      }
      
      if (filtered.length === 0 && untriedValues.length) {
        return helper(getRandomValue(incidents, field, alreadySelected), triedValues);
      }
      return filtered;
    };
    return helper(value === "RANDOM" ? getRandomValue(incidents, field, alreadySelected) : value, triedValues);
  },  
  
  filterIncidentsByFields = (incidents, fieldCriteria, agent, alreadySelected) => Object.entries(fieldCriteria).reduce((currentIncidents, [field, value]) => filterByCriterion(currentIncidents, field, value, agent, alreadySelected), incidents),
  selectIncidentsByConfiguration = (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) => { if (!Array.isArray(originalXlData) || !originalXlData.length) throw new Error("Invalid originalXlData"); const fieldToConfigKey = { "Service": "service", "Contact type": "contactType", "First time fix": "ftf" }; return Object.entries(sfAgentMapping).reduce((selectedIncidents, [sfMember, agents]) => { selectedIncidents[sfMember] = agents.reduce((agentIncidents, agent) => { const alreadySelected = new Set(); const usedServices = new Set(); agentIncidents[agent] = Array(maxIncidents).fill().reduce((incidents, _, i) => { const { order = ['service', 'contactType', 'ftf'], ...incidentConfig } = incidentConfigs[i % incidentConfigs.length]; const fieldCriteria = order.reduce((criteria, key) => { const field = Object.keys(fieldToConfigKey).find(field => fieldToConfigKey[field] === key); return field ? { ...criteria, [field]: incidentConfig[key] } : criteria; }, {}); let potentialIncidents = filterIncidentsByFields(originalXlData, fieldCriteria, agent, alreadySelected); if (potentialIncidents.length === 0 && fieldCriteria["Service"] && !usedServices.has(fieldCriteria["Service"])) { usedServices.add(fieldCriteria["Service"]); fieldCriteria["Service"] = getRandomValue(originalXlData, "Service", usedServices); potentialIncidents = filterIncidentsByFields(originalXlData, fieldCriteria, agent, alreadySelected); } const uniqueIncident = potentialIncidents.length && selectUniqueIncidentForAgent(potentialIncidents, alreadySelected); uniqueIncident && incidents.push(uniqueIncident); return incidents; }, []); let attempts = 0; 
  
  while (agentIncidents[agent].length < maxIncidents && attempts < maxIncidents) { const fieldCriteria = { "Service": getRandomValue(originalXlData, "Service", usedServices), "Contact type": getRandomValue(originalXlData, "Contact type", alreadySelected) }; 
  
  const potentialIncidents = filterIncidentsByFields(originalXlData, fieldCriteria, agent, alreadySelected); const uniqueIncident = potentialIncidents.length && selectUniqueIncidentForAgent(potentialIncidents, alreadySelected); if (uniqueIncident) { agentIncidents[agent].push(uniqueIncident); } else { attempts++; usedServices.add(fieldCriteria["Service"]); } } return agentIncidents; }, {}); return selectedIncidents; }, {}); }

} = {};

/* WRITE + DOWNLOAD */

const {

  createAndWriteWorksheet = (workbook, rows) => { const { json_to_sheet, book_append_sheet } = xlsx.utils; const { join } = path; const { SERV_FILENAME } = process.env; const newWorksheet = json_to_sheet(rows); const sheetName = "Processed List"; if (!workbook.Sheets[sheetName]) { book_append_sheet(workbook, newWorksheet, sheetName); } else { workbook.Sheets[sheetName] = newWorksheet; } const newFilePath = join(__dirname, "uploads", SERV_FILENAME); xlsx.writeFile(workbook, newFilePath); return newFilePath; },
  formatRowsForDownload = selectedIncidents => { let [previousSFMember, previousAgent] = ["", ""]; let rows = []; Object.entries(selectedIncidents).forEach(([sfMember, agents], sfIndex) => { Object.entries(agents).forEach(([agent, incidents], agentIndex) => { incidents.forEach(incident => { const row = { "SF Member": previousSFMember === sfMember ? "" : sfMember, Agent: previousAgent === agent ? "" : agent, ...["Task Number", "Service", "Contact type", "First time fix"].reduce((acc, key) => ({ ...acc, [key]: incident[key] }), {}) };[previousSFMember, previousAgent] = [sfMember, agent]; rows.push(row); }); agentIndex < Object.keys(agents).length - 1 ? rows.push({}) : null; }); sfIndex < Object.keys(selectedIncidents).length - 1 ? rows.push({}, {}) : null; }); return rows; },
  downloadFile = (res, newFilePath) => { const errorHandler = (err, message) => { if (err) throw new Error(`${message} ${err}`); }; const unlinkFile = (path, message) => fs.unlink(path, err => errorHandler(err, message)); res.download(newFilePath, filename, err => { errorHandler(err, "Error sending the file:"); unlinkFile(newFilePath, "Error deleting the processed file:"); unlinkFile(lastUploadedFilePath, "Error deleting the temporary file:"); }); }

} = {};