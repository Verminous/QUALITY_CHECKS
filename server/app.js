const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
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
  origin && allowedOrigins.includes(origin) ? res.setHeader('Access-Control-Allow-Origin', origin) : null;
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
    lastUploadedFilePath ? console.error("Last uploaded file path:", lastUploadedFilePath) : null;
    res.status(500).send("Internal Server Error");
  }
});

const createAndWriteWorksheet = (workbook, rows) => {
  const newWorksheet = xlsx.utils.json_to_sheet(rows);
  workbook.Sheets["Processed List"] ? workbook.Sheets["Processed List"] = newWorksheet : xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List");
  const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME);
  xlsx.writeFile(workbook, newFilePath);
  return newFilePath;
};

const downloadFile = (res, newFilePath) => {
  res.download(newFilePath, "AskIT - QCH_processed.xlsx", function (err) {
    if (err) throw new Error("Error sending the file: " + err);
  });
};

const mapIncidentsByAgent = (originalXlData) => {
  const incidentsByAgent = {};
  originalXlData.forEach((incident) => {
    const agent = incident["Taken By"];
    incidentsByAgent[agent] = incidentsByAgent[agent] ? incidentsByAgent[agent] : [];
    incidentsByAgent[agent].push(incident);
  });
  return incidentsByAgent;
}

const mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => {
  const sfAgentMapping = {};
  const shuffledAgents = Object.keys(incidentsByAgent).sort(() => 0.5 - Math.random());
  shuffledAgents.forEach((agent, index) => {
    const sfMember = sfMembers[index % sfMembers.length];
    sfAgentMapping[sfMember] = sfAgentMapping[sfMember] ? sfAgentMapping[sfMember] : [];
    sfAgentMapping[sfMember].push(agent);
  });
  return sfAgentMapping;
}

const formatRowsForDownload = (selectedIncidents) => {
  const rows = [];
  let previousSFMember = "";
  let previousAgent = "";
  Object.keys(selectedIncidents).forEach(sfMember => {
    Object.keys(selectedIncidents[sfMember]).forEach(agent => {
      selectedIncidents[sfMember][agent].forEach((incident) => {
        rows.push({
          "SF Member": previousSFMember === sfMember ? "" : sfMember,
          Agent: previousAgent === agent ? "" : agent,
          "Task Number": incident["Task Number"],
          Service: incident["Service"],
          "Contact type": incident["Contact type"],
          "First time fix": incident["First time fix"],
        });
        previousSFMember = previousSFMember !== sfMember ? sfMember : previousSFMember;
        previousAgent = previousAgent !== agent ? agent : previousAgent;
      });
    });
  });
  return rows;
};

const selectIncidentsByConfiguration = async (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) => {
  const selectedIncidents = {};
  const alreadySelected = {};
  Object.keys(sfAgentMapping).forEach(sfMember => {
    selectedIncidents[sfMember] = {};
    sfAgentMapping[sfMember].forEach(agent => {
      selectedIncidents[sfMember][agent] = [];
      alreadySelected[agent] = new Set();
      Array(maxIncidents).fill().forEach((_, i) => {
        const incidentConfig = incidentConfigs[i % incidentConfigs.length];
        let potentialIncidents = [...originalXlData];
        ['Service', 'Contact type', 'First time fix'].forEach(field => { potentialIncidents = filterIncidentsByCriterion(potentialIncidents, field, incidentConfig[field.toLowerCase()], agent, alreadySelected[agent]); });
        const selectedIncident = selectUniqueIncidentForAgent(potentialIncidents, alreadySelected[agent]);
        selectedIncident ? (selectedIncidents[sfMember][agent].push(selectedIncident), alreadySelected[agent].add(selectedIncident)) : null;
      });
    });
  });
  return selectedIncidents;
}

const filterIncidentsByCriterion = (incidents, field, value, agent, alreadySelected) => {
  value = (value === 'RANDOM') ? getRandomValue(incidents, field) : value;
  const filtered = incidents.filter(incident => !alreadySelected.has(incident) && incident[field] === value && incident['Taken By'] === agent);
  return filtered.length ? filtered : incidents.filter(incident => incident['Taken By'] === agent);
}

const getRandomValue = (incidents, field) => {
  const values = [...new Set(incidents.map(incident => incident[field]))];
  return values[Math.floor(Math.random() * values.length)];
}

const selectUniqueIncidentForAgent = (filteredIncidents, alreadySelected) => {
  const uniqueIncidents = filteredIncidents.filter(incident => !alreadySelected.has(incident));
  return uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : null;
};

app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});