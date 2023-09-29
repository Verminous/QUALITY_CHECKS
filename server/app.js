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
  const config = req.body;
  const workbook = xlsx.readFile(lastUploadedFilePath);
  const sheetName = workbook.SheetNames[0];
  const originalXlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
  const incidentsByAgent = {};
  const incidentConfigs = config.incidentConfigs;
  originalXlData.forEach((incident) => {
    const agent = incident["Taken By"];
    if (!incidentsByAgent[agent]) {
      incidentsByAgent[agent] = [];
    }
    incidentsByAgent[agent].push(incident);
  });
  const sfMembers = config.sfMembers;
  const agents = Object.keys(incidentsByAgent);
  const shuffledAgents = agents.sort(() => 0.5 - Math.random());
  const sfAgentMapping = {};
  shuffledAgents.forEach((agent, index) => {
    const sfMember = sfMembers[index % sfMembers.length];
    if (!sfAgentMapping[sfMember]) sfAgentMapping[sfMember] = [];
    sfAgentMapping[sfMember].push(agent);
  });
  const selectedIncidents = {};
  const maxIncidents = config.incidentsPerAgent || 10;
  const processedTaskNumbers = new Set();
  for (const sfMember in sfAgentMapping) {
    selectedIncidents[sfMember] = {};
    sfAgentMapping[sfMember].forEach((agent) => {
      selectedIncidents[sfMember][agent] = [];
      let incidentsForThisConfig = Math.ceil(
        maxIncidents / incidentConfigs.length
      );
      for (const incidentConfig of incidentConfigs) {
        let currentService =
          incidentConfig.service !== "RANDOM" ? incidentConfig.service : null;
        let currentContactType =
          incidentConfig.contactType !== "RANDOM"
            ? incidentConfig.contactType
            : null;
        let currentFtf =
          incidentConfig.ftf !== "RANDOM" ? incidentConfig.ftf : null;
        let potentialIncidents = incidentsByAgent[agent].filter((incident) => {
          const serviceMatches =
            incidentConfig.service !== "RANDOM"
              ? incident["Service"] === incidentConfig.service
              : true;
          const contactTypeMatches =
            incidentConfig.contactType !== "RANDOM"
              ? incident["Contact type"] === incidentConfig.contactType
              : true;
          const ftfMatches =
            incidentConfig.ftf !== "RANDOM"
              ? incident["First time fix"] == incidentConfig.ftf
              : true;
          return (
            !processedTaskNumbers.has(incident["Task Number"]) &&
            serviceMatches &&
            contactTypeMatches &&
            ftfMatches
          );
        });
        if (potentialIncidents.length < incidentsForThisConfig) {
          if (incidentConfig.service === "RANDOM")
            currentService = getRandomService();
          if (incidentConfig.contactType === "RANDOM")
            currentContactType = getRandomContactType();
          if (incidentConfig.ftf === "RANDOM") currentFtf = getRandomFtf();
          potentialIncidents = incidentsByAgent[agent].filter((incident) => {
            return (
              !processedTaskNumbers.has(incident["Task Number"]) &&
              currentService === incident["Service"] &&
              currentContactType === incident["Contact type"] &&
              currentFtf === incident["First time fix"]
            );
          });
        }
        const selectedFromThisConfig = potentialIncidents.slice(
          0,
          incidentsForThisConfig
        );
        selectedFromThisConfig.forEach((incident) => {
          processedTaskNumbers.add(incident["Task Number"]);
          selectedIncidents[sfMember][agent].push(incident);
        });
        incidentsForThisConfig =
          incidentsForThisConfig - selectedFromThisConfig.length;
      }
      let remainingIncidents =
        maxIncidents - selectedIncidents[sfMember][agent].length;
      for (const incidentConfig of incidentConfigs) {
        if (remainingIncidents <= 0) break;
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
            currentService === incident["Service"] &&
            currentContactType === incident["Contact type"] &&
            currentFtf === incident["First time fix"]
          );
        });
        const selectedForRemaining = potentialIncidents.slice(
          0,
          remainingIncidents
        );
        selectedForRemaining.forEach((incident) => {
          processedTaskNumbers.add(incident["Task Number"]);
          selectedIncidents[sfMember][agent].push(incident);
        });
        remainingIncidents = remainingIncidents - selectedForRemaining.length;
      }
    });
  }
  const rows = [];
  let previousSFMember = "";
  let previousAgent = "";
  for (const sfMember in selectedIncidents) {
    for (const agent in selectedIncidents[sfMember]) {
      for (
        let i = 0;
        i < incidentConfigs.length &&
        selectedIncidents[sfMember][agent].length < maxIncidents;
        i++
      ) {
        const incidentConfig = incidentConfigs[i];
        const filteredIncidents = originalXlData.filter((incident) => {
          return (
            (incidentConfig.service === "RANDOM" ||
              incidentConfig.service === incident["Service"]) &&
            (incidentConfig.contactType === "RANDOM" ||
              incidentConfig.contactType === incident["Contact type"]) &&
            (incidentConfig.ftf === "RANDOM" ||
              incidentConfig.ftf === incident["First time fix"]) &&
            !processedTaskNumbers.has(incident["Task Number"])
          );
        });
        filteredIncidents.forEach((incident) => {
          if (
            selectedIncidents[sfMember][agent].length < maxIncidents &&
            incident["Taken By"] === agent
          ) {
            processedTaskNumbers.add(incident["Task Number"]);
            selectedIncidents[sfMember][agent].push(incident);
          }
        });
      }
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
  const newWorkbook = workbook;
  if (rows.length < config.incidentsPerAgent) {
    console.error("Not enough incidents matched the provided configuration");
    res
      .status(500)
      .send("Not enough incidents matched the provided configuration");
    return;
  }
  const newWorksheet = xlsx.utils.json_to_sheet(rows);
  if (!newWorkbook.Sheets["Processed List"]) {
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Processed List");
  } else {
    newWorkbook.Sheets["Processed List"] = newWorksheet;
  }
  const newFilePath = path.join(
    __dirname,
    "uploads",
    process.env.SERV_FILENAME
  );
  try {
    xlsx.writeFile(newWorkbook, newFilePath);
  } catch (error) {
    console.error("Error writing the workbook:", error);
  }
  res.download(newFilePath, "AskIT - QCH_processed.xlsx", function (err) {
    if (err) {
      console.error("Error sending the file:", err);
    }
  });
});
app.listen(port, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
