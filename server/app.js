const express = require("express"), cors = require('cors'), fs = require('fs'), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
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

const {

  getRandomValue = (incidents, field, alreadySelected) => { const uniqueValues = [...new Set(incidents.map(i => i[field]))]; const unselectedValues = uniqueValues.filter(value => !alreadySelected.has(value)); const values = unselectedValues.length ? unselectedValues : uniqueValues; return values[Math.floor(Math.random() * values.length)]; },
  fisherYatesShuffle = array => array.reduceRight((acc, _, i, arr) => { const j = Math.floor(Math.random() * (i + 1)); [arr[i], arr[j]] = [arr[j], arr[i]]; return arr; }, array),
  mapIncidentsByAgent = (data) => data.reduce((acc, incident) => { const agent = incident["Taken By"]; acc[agent] = [...(acc[agent] || []), incident]; return acc; }, {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => { return fisherYatesShuffle(Object.keys(incidentsByAgent)).reduce((sfAgentMapping, agent, index) => { const sfMember = sfMembers[index % sfMembers.length]; sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent]; return sfAgentMapping; }, {}); },
  selectUniqueIncidentForAgent = (filteredIncidents = [], alreadySelected, originalIncidents = []) => { const getUniqueIncidents = incidents => incidents.filter(incident => !alreadySelected.has(incident)), getRandomIncident = incidents => incidents[Math.floor(Math.random() * incidents.length)], uniqueFilteredIncidents = getUniqueIncidents(filteredIncidents), selectedIncident = uniqueFilteredIncidents.length ? getRandomIncident(uniqueFilteredIncidents) : getRandomIncident(getUniqueIncidents(originalIncidents)); selectedIncident ? alreadySelected.add(selectedIncident) : console.warn("All original incidents have been selected for the current agent."); return selectedIncident; },

  filterByCriterion = ( incidents, field, value, agent, alreadySelected, triedValues = new Set() ) => { const helper = (value, triedValues) => { triedValues.add(value); const filtered = incidents.filter( (incident) => !alreadySelected.has(incident) && incident[field] === value && incident["Taken By"] === agent ); if (!filtered.length) { const untriedValues = incidents .map((i) => i[field]) .filter((value) => !triedValues.has(value)); return untriedValues.length ? helper(getRandomValue(incidents, field, alreadySelected), triedValues) : []; } return filtered; }; return helper( value === "RANDOM" ? getRandomValue(incidents, field, alreadySelected) : value, triedValues ); },
  filterIncidentsByFields = (incidents, fieldCriteria, agent, alreadySelected) => Object.entries(fieldCriteria).reduce( (currentIncidents, [field, value]) => filterByCriterion(currentIncidents, field, value, agent, alreadySelected), incidents ),
  selectIncidentsByConfiguration = (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) => {
    if (!Array.isArray(originalXlData) || !originalXlData.length) {
      throw new Error("Invalid originalXlData");
    }
    const fieldToConfigKey = { "Service": "service", "Contact type": "contactType", "First time fix": "ftf" };
  
    return Object.entries(sfAgentMapping).reduce((selectedIncidents, [sfMember, agents]) => {
      selectedIncidents[sfMember] = agents.reduce((agentIncidents, agent) => {
        const alreadySelected = new Set();
        agentIncidents[agent] = Array(maxIncidents).fill().reduce((incidents, _, i) => {
          const incidentConfig = incidentConfigs[i % incidentConfigs.length];
          let fields = Object.keys(incidentConfig).map(key => Object.keys(fieldToConfigKey).find(field => fieldToConfigKey[field] === key));
          let fieldCriteria = fields.reduce((criteria, field) => {
            const configKey = fieldToConfigKey[field];
            if (incidentConfig[configKey]) {
              criteria[field] = incidentConfig[configKey];
            }
            return criteria;
          }, {});
          const potentialIncidents = filterIncidentsByFields(originalXlData, fieldCriteria, agent, alreadySelected);
          potentialIncidents.length ? (uniqueIncident = selectUniqueIncidentForAgent(potentialIncidents, alreadySelected), incidents.push(uniqueIncident), alreadySelected.add(uniqueIncident.incidentId)) : console.log(`Warning: No incidents available for fallback for agent ${agent}.`);
          console.log(`Selected ${incidents.length} incidents for agent ${agent}`);
          return incidents;
        }, []);
        return agentIncidents;
      }, {});
      return selectedIncidents;
    }, {});
  }
} = {};

/* WRITE + DOWNLOAD */
const { createAndWriteWorksheet = (workbook, rows) => { const newWorksheet = xlsx.utils.json_to_sheet(rows); workbook.Sheets["Processed List"] ? (workbook.Sheets["Processed List"] = newWorksheet) : xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List"); const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME); xlsx.writeFile(workbook, newFilePath); return newFilePath; }, formatRowsForDownload = (selectedIncidents) => { let [previousSFMember, previousAgent] = [""]; return Object.keys(selectedIncidents).flatMap((sfMember) => Object.keys(selectedIncidents[sfMember]).flatMap((agent) => selectedIncidents[sfMember][agent].map((incident) => { const row = { "SF Member": previousSFMember === sfMember ? "" : sfMember, Agent: previousAgent === agent ? "" : agent, "Task Number": incident["Task Number"], Service: incident["Service"], "Contact type": incident["Contact type"], "First time fix": incident["First time fix"], };[previousSFMember, previousAgent] = [sfMember, agent]; return row; }))); }, downloadFile = (res, newFilePath) => { res.download(newFilePath, filename, (err) => { if (err) throw new Error("Error sending the file: " + err); fs.unlink(newFilePath, (err) => { if (err) throw new Error("Error deleting the processed file: " + err); console.log("Processed file deleted successfully"); }); fs.unlink(lastUploadedFilePath, (err) => { if (err) throw new Error("Error deleting the temporary file: " + err); console.log("Temporary file deleted successfully"); }); }); } } = {};