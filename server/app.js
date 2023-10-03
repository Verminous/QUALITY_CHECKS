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
app.use(cors());
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
function getRandomValue(options) {
  const index = Math.floor(Math.random() * options.length);
  return options[index];
}
function getRandomService() {
  return getRandomValue([
    "EMEIA Workplace",
    "Secure Internet Gateway (Global SIG)",
    "Identity and Access Management",
    "Identity Access Management (Finland)",
    "M365 Teams",
    "M365 Email",
    "M365 Apps",
    "Software Distribution (SCCM)",
    "Ask IT",
    "EMEIA Messaging",
    "Mobile Phones UK",
    "ZinZai Connect",
    "ForcePoint",
    "Network Service (CE/WEMEIA)",
    "M365 Sharepoint",
  ]);
}
function getRandomContactType() {
  return getRandomValue([
    "Self-service",
    "Phone - Unknown User",
    "Phone",
    "Chat",
  ]);
}
function getRandomFtf() {
  return getRandomValue([true, false]);
}
app.post("/process", upload.single("file"), async (req, res) => {
  try {
    const config = req.body;
    const workbook = xlsx.readFile(lastUploadedFilePath);
    const sheetName = workbook.SheetNames[0];
    const originalXlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    const incidentConfigs = config.incidentConfigs;
    const sfMembers = config.sfMembers;
    const incidentsByAgent = mapIncidentsByAgent(originalXlData);
    const sfAgentMapping = mapSFAgentsToIncidentAgents(
      sfMembers,
      incidentsByAgent
    );
    const selectedIncidents = await selectIncidentsByConfiguration(
      originalXlData,
      incidentConfigs,
      config.incidentsPerAgent,
      sfAgentMapping
    );
    const rows = formatRowsForDownload(selectedIncidents);
    if (rows.length < config.incidentsPerAgent) {
      console.error("Not enough incidents matched the provided configuration");
      return res
        .status(500)
        .send("Not enough incidents matched the provided configuration");
    }
    const newWorksheet = xlsx.utils.json_to_sheet(rows);
    if (!workbook.Sheets["Processed List"]) {
      xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List");
    } else {
      workbook.Sheets["Processed List"] = newWorksheet;
    }
    const newFilePath = path.join(
      __dirname,
      "uploads",
      process.env.SERV_FILENAME
    );
    try {
      xlsx.writeFile(workbook, newFilePath);
    } catch (error) {
      console.error("Error writing the workbook:", error);
    }
    res.download(newFilePath, "AskIT - QCH_processed.xlsx", function (err) {
      if (err) {
        console.error("Error sending the file:", err);
      }
    });
  } catch (error) {
    console.error("Error in /process:", error);
    console.error("Request body:", req.body);
    if (lastUploadedFilePath) {
      console.error("Last uploaded file path:", lastUploadedFilePath);
    }
    res.status(500).send("Internal Server Error");
  }
});
function mapIncidentsByAgent(originalXlData) {
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
function mapSFAgentsToIncidentAgents(sfMembers, incidentsByAgent) {
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
function selectIncidentsByConfiguration(
  originalXlData,
  incidentConfigs,
  maxIncidents,
  sfAgentMapping
) {
  const incidentsByAgent = {};
  const selectedIncidents = {};
  const processedTaskNumbersByAgent = {}; 
  const processedTaskNumbers = new Set();
  originalXlData.forEach((incident) => {
    const agent = incident["Taken By"];
    if (!incidentsByAgent[agent]) {
      incidentsByAgent[agent] = [];
    }
    incidentsByAgent[agent].push(incident);
  });
  for (const sfMember in sfAgentMapping) {
    selectedIncidents[sfMember] = {};
    sfAgentMapping[sfMember].forEach((agent) => {
      if (!processedTaskNumbersByAgent[agent]) { 
        processedTaskNumbersByAgent[agent] = new Set();
      }
      
      selectedIncidents[sfMember][agent] = [];
      for (const incidentConfig of incidentConfigs) {
        let currentService =
          incidentConfig.service !== "RANDOM"
            ? incidentConfig.service
            : getRandomService();
        let currentContactType =
          incidentConfig.contactType !== "RANDOM"
            ? incidentConfig.contactType
            : getRandomContactType();
        let currentFtf =
          incidentConfig.ftf !== "RANDOM" ? incidentConfig.ftf : getRandomFtf();
          let potentialIncidents = incidentsByAgent[agent].filter((incident) => {
            return (
              !processedTaskNumbers.has(incident["Task Number"]) && 
              !processedTaskNumbersByAgent[agent].has(incident["Task Number"]) && 
              currentService === incident["Service"] &&
              currentContactType === incident["Contact type"] &&
              currentFtf === incident["First time fix"]
            );
          });
        let incidentsToAssign = Math.min(
          potentialIncidents.length,
          maxIncidents - selectedIncidents[sfMember][agent].length
        );
        for (let i = 0; i < incidentsToAssign; i++) {
          let incident = potentialIncidents[i];
          selectedIncidents[sfMember][agent].push(incident);
          processedTaskNumbers.add(incident["Task Number"]); 
          processedTaskNumbersByAgent[agent].add(incident["Task Number"]);
        }
      }
    });
  }
  return selectedIncidents;
}
function formatRowsForDownload(selectedIncidents) {
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
}
app.listen(port, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
