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

  if (origin && allowedOrigins.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  }

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

const getRandomValue = (options) => {
  const index = Math.floor(Math.random() * options.length);
  return options[index];
}

const getRandomService = () => {
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

const getRandomContactType = () => {
  return getRandomValue([
    "Self-service",
    "Phone - Unknown User",
    "Phone",
    "Chat",
  ]);
}

const getRandomFtf = () => {
  return getRandomValue([true, false]);
}

app.post("/process", upload.single("file"), async (req, res) => {
  try {
    const config = req.body; /**/ console.log(config);
    const workbook = xlsx.readFile(lastUploadedFilePath);
    const sheetName = workbook.SheetNames[0];
    const originalXlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    const incidentConfigs = config.incidentConfigs;
    const sfMembers = config.sfMembers;
    const agentNames = req.body.agentNames; // Get the agent names from the request body
    const incidentsByAgent = mapIncidentsByAgent(originalXlData, agentNames); // Pass the agent names to the function
    const sfAgentMapping = mapSFMembersToIncidentAgents(
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

const mapIncidentsByAgent = (originalXlData, agentNames) => { // Add a new parameter for the agent names
  const incidentsByAgent = {};
  originalXlData.forEach((incident) => {
    const agent = incident["Taken By"];
    if (agentNames.includes(agent)) { // Only consider incidents by agents in the provided list
      if (!incidentsByAgent[agent]) {
        incidentsByAgent[agent] = [];
      }
      incidentsByAgent[agent].push(incident);
    }
  });
  return incidentsByAgent;
};

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

const selectIncidentsByConfiguration = async (
  originalXlData,
  incidentConfigs,
  maxIncidents,
  sfAgentMapping
) => {
  const selectedIncidents = {};
  const processedTaskNumbersByAgent = {};
  const processedTaskNumbers = new Set();

  for (const sfMember in sfAgentMapping) {
    selectedIncidents[sfMember] = {};

    sfAgentMapping[sfMember].forEach((agent) => {
      if (!processedTaskNumbersByAgent[agent]) {
        processedTaskNumbersByAgent[agent] = new Set();
      }

      selectedIncidents[sfMember][agent] = [];

      for (const incidentConfig of incidentConfigs) {
        const filteredIncidents = originalXlData.filter((incident) => {
          return (
            !processedTaskNumbers.has(incident["Task Number"]) &&
            !processedTaskNumbersByAgent[agent].has(incident["Task Number"]) &&
            incident["Taken By"] === agent &&
            (incidentConfig.service === "RANDOM" || incidentConfig.service === incident["Service"]) &&
            (incidentConfig.contactType === "RANDOM" || incidentConfig.contactType === incident["Contact type"]) &&
            (incidentConfig.ftf === "RANDOM" || incidentConfig.ftf === incident["First time fix"])
          );
        });

        const toBeSelected = Math.min(
          filteredIncidents.length,
          maxIncidents - selectedIncidents[sfMember][agent].length
        );

        for (let i = 0; i < toBeSelected; i++) {
          const incident = filteredIncidents[i];
          selectedIncidents[sfMember][agent].push(incident);
          processedTaskNumbers.add(incident["Task Number"]);
          processedTaskNumbersByAgent[agent].add(incident["Task Number"]);
        }
      }
    });
  }
  return selectedIncidents;
};

app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});

