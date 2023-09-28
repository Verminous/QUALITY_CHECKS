const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
require('dotenv').config({ path: path.join(__dirname, '.env') });

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

app.post("/process", upload.single("file"), async (req, res) => {
    const config = req.body;
    const workbook = xlsx.readFile(lastUploadedFilePath);
    const sheetName = workbook.SheetNames[0];
    const xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const incidentsByAgent = {};
    const incidentConfigs = config.incidentConfigs;
    const getIncidentConfig = (index) => incidentConfigs[index % incidentConfigs.length];

    xlData.forEach((incident, index) => {
        const agent = incident["Taken By"];
        if (!incidentsByAgent[agent]) {
            incidentsByAgent[agent] = [];
        }
        const incidentConfig = getIncidentConfig(incidentsByAgent[agent].length);
        incident["Service"] = incidentConfig.service === 'RANDOM' ? incident["Service"] : incidentConfig.service;
        incident["Contact type"] = incidentConfig.contactType === 'RANDOM' ? incident["Contact type"] : incidentConfig.contactType;
        incident["First time fix"] = incidentConfig.ftf === 'RANDOM' ? incident["First time fix"] : incidentConfig.ftf;
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

    for (const sfMember in sfAgentMapping) {
        selectedIncidents[sfMember] = {};
        sfAgentMapping[sfMember].forEach((agent) => {
            const agentIncidents = incidentsByAgent[agent].slice(0, maxIncidents);
            selectedIncidents[sfMember][agent] = agentIncidents;
        });
    }

    const rows = [];
    let previousSFMember = "";
    let previousAgent = "";
    for (const sfMember in selectedIncidents) {
        for (const agent in selectedIncidents[sfMember]) {
            selectedIncidents[sfMember][agent].forEach((incident, index) => {
                rows.push({
                    "SF Member": previousSFMember === sfMember ? "" : sfMember,
                    Agent: previousAgent === agent ? "" : agent,
                    "Task Number": incident["Task Number"],
                    Service: incident["Service"],
                    "Contact Type": incident["Contact type"],
                    "First Time Fix": incident["First time fix"],
                });
                if (previousSFMember !== sfMember) previousSFMember = sfMember;
                if (previousAgent !== agent) previousAgent = agent;
            });
        }
    }

    const newWorkbook = workbook;

    if (rows.length === 0) {
        res.status(500).send('No incidents matched the provided configuration');
        return;
    }

    const newWorksheet = xlsx.utils.json_to_sheet(rows);

    if (!newWorkbook.Sheets["Processed List"]) {
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Processed List");
    } else {
        newWorkbook.Sheets["Processed List"] = newWorksheet;
    }

    const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME);

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