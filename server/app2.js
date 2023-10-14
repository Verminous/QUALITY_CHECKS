const express = require("express"), cors = require('cors'),  fs = require('fs'), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(cors());
app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => { const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`]; allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization')); next(); });
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => { console.log(`Server running at http://${hostname}:${port}/`); });

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => { lastUploadedFilePath = p; const workbook = xlsx.readFile(p), sheet = workbook.Sheets[workbook.SheetNames[0]], json = xlsx.utils.sheet_to_json(sheet); res.json({ agentNames: [...new Set(json.map(data => data["Taken By"]))] }); });
app.post("/process", upload.single("file"), async ({ body: config }, res) => { console.log(config); try { const workbook = xlsx.readFile(lastUploadedFilePath), sheet = workbook.Sheets[workbook.SheetNames[0]], originalXlData = xlsx.utils.sheet_to_json(sheet), { incidentConfigs, sfMembers, incidentsPerAgent } = config, selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, incidentsPerAgent, mapSFMembersToIncidentAgents(sfMembers, mapIncidentsByAgent(originalXlData))), rows = formatRowsForDownload(selectedIncidents); if (rows.length < incidentsPerAgent) { throw new Error("Not enough incidents matched the provided configuration"); } else { downloadFile(res, createAndWriteWorksheet(workbook, rows)); }; console.log(config); } catch (error) { console.error("Error in /process:", error, "Request body:", config, lastUploadedFilePath && "Last uploaded file path:", lastUploadedFilePath); res.status(500).send("Internal Server Error"); } });

const logDir = path.join(__dirname, 'logs'); if (!fs.existsSync(logDir)) { fs.mkdirSync(logDir); } function logToFile(message) { console.log(message); fs.appendFile(path.join(logDir, 'log.txt'), message + '\n', (err) => { if (err) throw err; }); }

const {

  getRandomValue = (incidents, field, alreadySelected) => { const uniqueValues = [...new Set(incidents.map((i) => i[field]))]; const unselectedValues = uniqueValues.filter((value) => !alreadySelected.has(value)); const valuesToUse = unselectedValues.length > 0 ? unselectedValues : uniqueValues; return valuesToUse[Math.floor(Math.random() * valuesToUse.length)]; },
  fisherYatesShuffle = (array) => { array.forEach((_, i) => { const j = Math.floor(Math.random() * (i + 1));[array[i], array[j]] = [array[j], array[i]]; }); return array; },
  mapIncidentsByAgent = (data) => data.reduce((acc, incident) => { const agent = incident["Taken By"]; acc[agent] = acc[agent] || []; acc[agent].push(incident); return acc; }, {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => { const sfAgentMapping = {}, agents = Object.keys(incidentsByAgent); fisherYatesShuffle(agents).forEach((agent, index) => { const sfMember = sfMembers[index % sfMembers.length]; sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent]; }); return sfAgentMapping; },
  selectUniqueIncident = (filteredIncidents, alreadySelected) => { const uniqueIncidents = filteredIncidents.filter((incident) => !alreadySelected.has(incident)); return uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : null; },
  selectUniqueIncidentForAgent = (filteredIncidents = [], alreadySelected, originalIncidents = []) => { let uniqueIncidents = filteredIncidents.filter((incident) => !alreadySelected.has(incident)); let selectedIncident; selectedIncident = uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : (() => { let remainingOriginals = originalIncidents.filter((incident) => !alreadySelected.has(incident)); return remainingOriginals.length ? remainingOriginals[Math.floor(Math.random() * remainingOriginals.length)] : (console.warn('All original incidents have been selected for the current agent.'), null); })(); selectedIncident ? alreadySelected.add(selectedIncident) : null; return selectedIncident; },

  filterByCriterion = (incidents, field, value, agent, alreadySelected, triedValues = new Set()) => {
    value = value === "RANDOM" ? getRandomValue(incidents, field, alreadySelected) : value;
    triedValues.add(value);
    const filtered = incidents.filter(incident => {
      const matches = !alreadySelected.has(incident) && incident[field] === value && incident["Taken By"] === agent;
      return matches;
    });
  
    if (!filtered.length) {
      const allValues = incidents.map((i) => i[field]);
      const untriedValues = allValues.filter((value) => !triedValues.has(value));
      if (untriedValues.length === 0) {
        return [];
      }
      value = getRandomValue(incidents, field, alreadySelected);
      return filterByCriterion(incidents, field, value, agent, alreadySelected, triedValues);
    }
  
    return filtered;
  },

  filterIncidentsByCriterion = (incidents, field, value, agent, alreadySelected, originalIncidents) => {
    const filtered = filterByCriterion(incidents, field, value, agent, alreadySelected);
    logToFile(`filterIncidentsByCriterion_1 - Filtered ${filtered.length} incidents by ${field} = ${value} for agent ${agent}`);
    const selectedIncident = selectUniqueIncident(filtered, alreadySelected);
    logToFile(`filterIncidentsByCriterion_2 - Selected incident after filtering by ${field} = ${value} for agent ${agent}: ${JSON.stringify(selectedIncident)}`);    if (!selectedIncident && value !== 'RANDOM') {
      const fallbackFiltered = originalIncidents.filter(incident => incident[field] === value && !alreadySelected.has(incident));
      return fallbackFiltered.length ? [fallbackFiltered[Math.floor(Math.random() * fallbackFiltered.length)]] : [];
    }
    logToFile(`filterIncidentsByCriterion_3 - Selected ${selectedIncident ? 1 : 0} incidents after filtering by ${field} = ${value} for agent ${agent}`);
    return selectedIncident ? [selectedIncident] : [];
  },

   filterIncidentsByFields = (incidents, fieldCriteria, agent, alreadySelected) => {
    return Object.entries(fieldCriteria).reduce((currentIncidents, [field, value]) => {
        return filterByCriterion(currentIncidents, field, value, agent, alreadySelected);
    }, incidents);
},

 selectIncidentsByConfiguration = (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) => {
    (!Array.isArray(originalXlData) || !originalXlData.length) && (() => {
        throw new Error("Invalid originalXlData");
    })();

    const fieldToConfigKey = {
        "Service": "service",
        "Contact type": "contactType",
        "First time fix": "ftf"
    };

    return Object.entries(sfAgentMapping).reduce((selectedIncidents, [sfMember, agents]) => {
        selectedIncidents[sfMember] = agents.reduce((agentIncidents, agent) => {
            const alreadySelected = new Set();
            agentIncidents[agent] = Array(maxIncidents).fill().reduce((incidents, _, i) => {
                const incidentConfig = incidentConfigs[i % incidentConfigs.length];
                let fieldCriteria = {};
                ["Service", "Contact type", "First time fix"].forEach(field => {
                    const configKey = fieldToConfigKey[field];
                    if (incidentConfig[configKey]) {
                        fieldCriteria[field] = incidentConfig[configKey];
                    }
                });

                const potentialIncidents = filterIncidentsByFields(originalXlData, fieldCriteria, agent, alreadySelected);
                if (potentialIncidents.length) {
                    const uniqueIncident = selectUniqueIncidentForAgent(potentialIncidents, alreadySelected);
                    incidents.push(uniqueIncident);
                    alreadySelected.add(uniqueIncident.incidentId);
                } else {
                    logToFile(`Warning: No incidents available for fallback for agent ${agent}.`);
                }
                logToFile(`selectIncidentsByConfiguration_2 - Selected ${incidents.length} incidents for agent ${agent}`);
                return incidents;
            }, []);
            return agentIncidents;
        }, {});
        return selectedIncidents;
    }, {});
},
} = {};

/* WRITE + DOWNLOAD */
const { createAndWriteWorksheet = (workbook, rows) => { const newWorksheet = xlsx.utils.json_to_sheet(rows); workbook.Sheets["Processed List"] ? (workbook.Sheets["Processed List"] = newWorksheet) : xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List"); const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME); xlsx.writeFile(workbook, newFilePath); return newFilePath; }, formatRowsForDownload = (selectedIncidents) => { let [previousSFMember, previousAgent] = [""]; return Object.keys(selectedIncidents).flatMap((sfMember) => Object.keys(selectedIncidents[sfMember]).flatMap((agent) => selectedIncidents[sfMember][agent].map((incident) => { const row = { "SF Member": previousSFMember === sfMember ? "" : sfMember, Agent: previousAgent === agent ? "" : agent, "Task Number": incident["Task Number"], Service: incident["Service"], "Contact type": incident["Contact type"], "First time fix": incident["First time fix"], };[previousSFMember, previousAgent] = [sfMember, agent]; return row; }))); }, downloadFile = (res, newFilePath) => { res.download(newFilePath, filename, (err) => { if (err) throw new Error("Error sending the file: " + err); fs.unlink(newFilePath, (err) => { if (err) throw new Error("Error deleting the processed file: " + err); console.log("Processed file deleted successfully"); }); fs.unlink(lastUploadedFilePath, (err) => { if (err) throw new Error("Error deleting the temporary file: " + err); console.log("Temporary file deleted successfully"); }); }); } } = {};