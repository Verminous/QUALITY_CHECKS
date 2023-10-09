const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
require("dotenv").config({ path: path.join(__dirname, ".env") });
const app = express();
const port = process.env.SERV_PORT;
const hostname = process.env.HOSTNAME;
const upload = multer({ dest: "uploads/" });

app.use(bodyParser.json());

app.use((req, res, next) => {
  const allowedOrigins = ['http://localhost:8080', `http://${req.hostname}:8080`];
  const origin = req.headers.origin;
  if (origin && allowedOrigins.includes(origin)) { res.setHeader('Access-Control-Allow-Origin', origin); }
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('Access-Control-Allow-Credentials', true);
  next();
});

app.get("/", (req, res) => {
  res.send("Server running!");
});

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), (req, res) => {
  lastUploadedFilePath = req.file.path;
  const workbook = xlsx.readFile(lastUploadedFilePath);
  const sheet_name_list = workbook.SheetNames;
  const xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  const agentNames = [...new Set(xlData.map((data) => data["Taken By"]))];
  res.json({ agentNames });
});

const createAndWriteWorksheet = (workbook, rows) => {
  const newWorksheet = xlsx.utils.json_to_sheet(rows);
  if (!workbook.Sheets["Processed List"]) {
    xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List");
  } else {
    workbook.Sheets["Processed List"] = newWorksheet;
  }
  const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME);
  xlsx.writeFile(workbook, newFilePath);
  return newFilePath;
};

const downloadFile = (res, newFilePath) => {
  res.download(newFilePath, "AskIT - QCH_processed.xlsx", function (err) {
    if (err) throw new Error("Error sending the file: " + err);
  });
};

app.post("/process", upload.single("file"), async (req, res) => {
  try {
    const config = req.body;
    const workbook = xlsx.readFile(lastUploadedFilePath);
    const sheetName = workbook.SheetNames[0];
    const originalXlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    const incidentConfigs = config.incidentConfigs;
    const sfMembers = config.sfMembers;
    const incidentsByAgent = mapIncidentsByAgent(originalXlData);
    const sfAgentMapping = mapSFMembersToIncidentAgents(sfMembers, incidentsByAgent);
    const selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, config.incidentsPerAgent, sfAgentMapping);
    const rows = formatRowsForDownload(selectedIncidents);
    if (rows.length < config.incidentsPerAgent) throw new Error("Not enough incidents matched the provided configuration");
    const newFilePath = createAndWriteWorksheet(workbook, rows);
    downloadFile(res, newFilePath);
  } catch (error) {
    console.error("Error in /process:", error);
    console.error("Request body:", req.body);
    if (lastUploadedFilePath) console.error("Last uploaded file path:", lastUploadedFilePath);
    res.status(500).send("Internal Server Error");
  }
});

const mapIncidentsByAgent = (originalXlData) => {
  const incidentsByAgent = {};
  originalXlData.forEach((incident) => {
    const agent = incident["Taken By"];
    if (!incidentsByAgent[agent]) {
      incidentsByAgent[agent] = [];
    }
    incidentsByAgent[agent].push(incident);
  });
  return incidentsByAgent;
}

const mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => {
  const sfAgentMapping = {};
  const shuffledAgents = Object.keys(incidentsByAgent).sort(
    () => 0.5 - Math.random()
  );
  shuffledAgents.forEach((agent, index) => {
    const sfMember = sfMembers[index % sfMembers.length];
    if (!sfAgentMapping[sfMember]) sfAgentMapping[sfMember] = [];
    sfAgentMapping[sfMember].push(agent);
  });
  return sfAgentMapping;
}

const formatRowsForDownload = (selectedIncidents) => {
  const rows = [];
  let previousSFMember = "";
  let previousAgent = "";
  for (const sfMember in selectedIncidents) {
    for (const agent in selectedIncidents[sfMember]) {
      selectedIncidents[sfMember][agent].forEach((incident) => {
        rows.push({
          "SF Member": previousSFMember === sfMember ? "" : sfMember,
          Agent: previousAgent === agent ? "" : agent,
          "Task Number": incident["Task Number"],
          Service: incident["Service"],
          "Contact type": incident["Contact type"],
          "First time fix": incident["First time fix"],
        });
        if (previousSFMember !== sfMember) previousSFMember = sfMember;
        if (previousAgent !== agent) previousAgent = agent;
      });
    }
  }
  return rows;
};

function filterIncidentsByCriterion(incidents, field, value, agent) {
  if (value === 'RANDOM') {
      const uniqueValues = [...new Set(incidents.map(incident => incident[field]))];
      value = uniqueValues.length ? uniqueValues[Math.floor(Math.random() * uniqueValues.length)] : value;
  }
  
  const filtered = incidents.filter(incident => incident[field] === value && incident['Taken By'] === agent);
  
  if (!filtered.length) {
      return incidents.filter(incident => incident['Taken By'] === agent);
  }

  return filtered;
}

const selectUniqueIncidentForAgent = (filteredIncidents, processedTaskNumbers, agentTaskNumbers) => {
  const unassignedIncidents = filteredIncidents.filter(incident =>
      !agentTaskNumbers.has(incident['Task Number'])
  );
  return unassignedIncidents.length ? unassignedIncidents[Math.floor(Math.random() * unassignedIncidents.length)] : null;
};

async function selectIncidentsByConfiguration(originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) {
  const selectedIncidents = {};
  const processedTaskNumbers = new Set();
  const processedTaskNumbersByAgent = {};

  for (const sfMember in sfAgentMapping) {
      selectedIncidents[sfMember] = {};
      sfAgentMapping[sfMember].forEach(agent => {
          if (!processedTaskNumbersByAgent[agent]) {
              processedTaskNumbersByAgent[agent] = new Set();
          }
          selectedIncidents[sfMember][agent] = [];
          for (const incidentConfig of incidentConfigs) {
              let potentialIncidents = [...originalXlData];

              potentialIncidents = filterIncidentsByCriterion(potentialIncidents, 'Service', incidentConfig.service, agent);
              potentialIncidents = filterIncidentsByCriterion(potentialIncidents, 'Contact type', incidentConfig.contactType, agent);
              potentialIncidents = filterIncidentsByCriterion(potentialIncidents, 'First time fix', incidentConfig.ftf, agent);

              const selectedIncident = selectUniqueIncidentForAgent(potentialIncidents, processedTaskNumbers, processedTaskNumbersByAgent[agent]);

              if (selectedIncident) {
                  selectedIncidents[sfMember][agent].push(selectedIncident);
                  processedTaskNumbersByAgent[agent].add(selectedIncident['Task Number']);
              }
          }
      });
  }
  return selectedIncidents;
}


app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});