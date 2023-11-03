const express = require("express"), cors = require('cors'), fs = require('fs'), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(cors());
app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => { const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`]; allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization')); next(); });
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => { console.log(`Server running at http://${hostname}:${port}/`); });

/* LOGGER */

const logDir = path.join(__dirname, 'logs'); if (!fs.existsSync(logDir)) { fs.mkdirSync(logDir); } function logToFile(message) { console.log(message); fs.appendFile(path.join(logDir, 'log.txt'), message + '\n', (err) => { if (err) throw err; }); }

/* ENDPOINTS */

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => { lastUploadedFilePath = p; const workbook = xlsx.readFile(p), sheet = workbook.Sheets[workbook.SheetNames[0]], json = xlsx.utils.sheet_to_json(sheet); res.json({ agentNames: [...new Set(json.map(data => data["Taken By"]))] }); });
app.post("/process", upload.single("file"), async ({ body: config }, res) => {
  console.log(config);
  try {
      const workbook = xlsx.readFile(lastUploadedFilePath),
          sheet = workbook.Sheets[workbook.SheetNames[0]],
          originalXlData = xlsx.utils.sheet_to_json(sheet),
          { incidentConfigs, sfMembers, incidentsPerAgent } = config;

      let randomServices = config.randomServices;
      if (typeof randomServices === 'string') {
          randomServices = randomServices.split(',');
      }

      console.log(`randomServices before selectIncidentsByConfiguration: ${randomServices}`);
      const selectedIncidents = await selectIncidentsByConfiguration(
          originalXlData,
          incidentConfigs,
          incidentsPerAgent,
          randomServices,
          mapSFMembersToIncidentAgents(sfMembers, mapIncidentsByAgent(originalXlData)),
      );

      const rows = formatRowsForDownload(selectedIncidents);
      if (rows.length < incidentsPerAgent) {
          throw new Error("Not enough incidents matched the provided configuration");
      } else {
          downloadFile(res, createAndWriteWorksheet(workbook, rows));
      }
      console.log(config);
  } catch (error) {
      console.error("Error in /process:", error, "Request body:", config, lastUploadedFilePath && "Last uploaded file path:", lastUploadedFilePath);
      res.status(500).send("Internal Server Error");
  }
});

/* DATA PROCESSING */

const {

  getRandomValue = (incidents, field, alreadySelected) => {
    const uniqueValues = [...new Set(incidents.map((i) => i[field]))];
    const unselectedValues = uniqueValues.filter((value) => !alreadySelected.has(value));
    const valuesToUse = unselectedValues.length > 0 ? unselectedValues : uniqueValues;
    return valuesToUse[Math.floor(Math.random() * valuesToUse.length)];
  },

  getRandomService = (incidents, randomServices, alreadySelected) => {
    if (!Array.isArray(randomServices)) {
        throw new Error("randomServices is not an array");
    }
    if (randomServices.length === 0) {
        throw new Error("randomServices list is empty");
    }
    let servicesToTry = [...randomServices];
    while (servicesToTry.length) {
        const serviceIndex = Math.floor(Math.random() * servicesToTry.length);
        const service = servicesToTry[serviceIndex];
        const matchingIncidents = incidents.filter(incident => incident['Service'] === service && !alreadySelected.has(incident));
        if (matchingIncidents.length) {
            return service;
        } else {
            servicesToTry.splice(serviceIndex, 1);
        }
    }
    // Fallback logic
    const uniqueValues = [...new Set(incidents.map((i) => i['Service']))];
    const unselectedValues = uniqueValues.filter((value) => !alreadySelected.has(value));
    const valuesToUse = unselectedValues.length > 0 ? unselectedValues : uniqueValues;
    return valuesToUse[Math.floor(Math.random() * valuesToUse.length)];
},

  fisherYatesShuffle = (array) => { array.forEach((_, i) => { const j = Math.floor(Math.random() * (i + 1));[array[i], array[j]] = [array[j], array[i]]; }); return array; },
  mapIncidentsByAgent = (data) => data.reduce((acc, incident) => { const agent = incident["Taken By"]; acc[agent] = acc[agent] || []; acc[agent].push(incident); return acc; }, {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => { const sfAgentMapping = {}, agents = Object.keys(incidentsByAgent); fisherYatesShuffle(agents).forEach((agent, index) => { const sfMember = sfMembers[index % sfMembers.length]; sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent]; }); return sfAgentMapping; },
  selectUniqueIncidentForAgent = (filteredIncidents = [], alreadySelected, originalIncidents = []) => { let uniqueIncidents = filteredIncidents.filter((incident) => !alreadySelected.has(incident)); let selectedIncident; selectedIncident = uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : (() => { let remainingOriginals = originalIncidents.filter((incident) => !alreadySelected.has(incident)); return remainingOriginals.length ? remainingOriginals[Math.floor(Math.random() * remainingOriginals.length)] : (console.warn('All original incidents have been selected for the current agent.'), null); })(); selectedIncident ? alreadySelected.add(selectedIncident) : null; return selectedIncident; },

  filterByCriterion = (incidents, field, value, agent, alreadySelected, randomServices, triedValues = new Set()) => {
    if (!Array.isArray(incidents)) {
      console.error('Invalid incidents data. Expected an array.');
      return [];
    }
    if (field === 'Service' && value === 'RANDOM') {
      value = getRandomService(incidents, randomServices, alreadySelected);
    } else if (value === 'RANDOM') {
      value = getRandomValue(incidents, field, alreadySelected);
    }
    const filtered = incidents.filter(incident => {
      const matches = !alreadySelected.has(incident) && incident[field] === value && incident["Taken By"] === agent;
      return matches;
    });

    triedValues.add(value);

    if (!filtered.length) {
      const allValues = incidents.map((i) => i[field]);
      const untriedValues = allValues.filter((value) => !triedValues.has(value));

      if (untriedValues.length === 0) {
        return [];
      }

      value = getRandomValue(incidents, field, alreadySelected);
      return filterByCriterion(incidents, field, value, agent, alreadySelected, randomServices, triedValues);
    }
    return filtered;
  },

  filterIncidentsByFields = (incidents, fieldCriteria, agent, alreadySelected, randomServices) => {
    return Object.entries(fieldCriteria).reduce((currentIncidents, [field, value]) => {
      return filterByCriterion(currentIncidents, field, value, agent, alreadySelected, randomServices, new Set());
    }, incidents);
  },
  selectIncidentsByConfiguration = (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping, randomServices) => {
    if (!Array.isArray(originalXlData) || !originalXlData.length) {
      throw new Error("Invalid originalXlData");
    }

    const fieldToConfigKey = {
      "Service": "service",
      "Contact type": "contactType",
      "First time fix": "ftf"
    };

    let selectedIncidents = {};

    for (const [sfMember, agents] of Object.entries(sfAgentMapping)) {
      selectedIncidents[sfMember] = {};

      for (const agent of agents) {
        const alreadySelected = new Set();
        selectedIncidents[sfMember][agent] = Array(maxIncidents).fill().reduce((incidents, _, i) => {
          const incidentConfig = incidentConfigs[i % incidentConfigs.length];
          let fieldCriteria = {};

          ["Service", "Contact type", "First time fix"].forEach(field => {
            const configKey = fieldToConfigKey[field];
            if (incidentConfig[configKey]) {
              fieldCriteria[field] = incidentConfig[configKey];
            }
          });

          const potentialIncidents = filterIncidentsByFields(originalXlData, fieldCriteria, agent, alreadySelected, randomServices);

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
      }
    }

    return selectedIncidents;
  }
} = {};

/* WRITE + DOWNLOAD */
const { createAndWriteWorksheet = (workbook, rows) => { const { json_to_sheet, book_append_sheet } = xlsx.utils; const { join } = path; const { SERV_FILENAME } = process.env; const newWorksheet = json_to_sheet(rows); const sheetName = "Processed List"; if (!workbook.Sheets[sheetName]) { book_append_sheet(workbook, newWorksheet, sheetName); } else { workbook.Sheets[sheetName] = newWorksheet; } const newFilePath = join(__dirname, "uploads", SERV_FILENAME); xlsx.writeFile(workbook, newFilePath); return newFilePath; }, formatRowsForDownload = selectedIncidents => { let [previousSFMember, previousAgent] = ["", ""]; let rows = []; Object.entries(selectedIncidents).forEach(([sfMember, agents], sfIndex) => { Object.entries(agents).forEach(([agent, incidents], agentIndex) => { incidents.forEach(incident => { const row = { "SF Member": previousSFMember === sfMember ? "" : sfMember, Agent: previousAgent === agent ? "" : agent, ...["Task Number", "Service", "Contact type", "First time fix"].reduce((acc, key) => ({ ...acc, [key]: incident[key] }), {}) };[previousSFMember, previousAgent] = [sfMember, agent]; rows.push(row); }); agentIndex < Object.keys(agents).length - 1 ? rows.push({}) : null; }); sfIndex < Object.keys(selectedIncidents).length - 1 ? rows.push({}, {}) : null; }); return rows; }, downloadFile = (res, newFilePath) => { const errorHandler = (err, message) => { if (err) throw new Error(`${message} ${err}`); }; const unlinkFile = (path, message) => fs.unlink(path, err => errorHandler(err, message)); res.download(newFilePath, filename, err => { errorHandler(err, "Error sending the file:"); unlinkFile(newFilePath, "Error deleting the processed file:"); unlinkFile(lastUploadedFilePath, "Error deleting the temporary file:"); }); } } = {};