const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const path = require("path");

require("dotenv").config({
    path: path.join(__dirname, ".env"),
});

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
            
            let incidentsForThisConfig = Math.ceil(maxIncidents / incidentConfigs.length);
            
            incidentConfigs.forEach((incidentConfig) => {
                const potentialIncidents = originalXlData.filter(incident => {
                    return (!processedTaskNumbers.has(incident["Task Number"])) &&
                           (incidentConfig.service === "RANDOM" || incidentConfig.service === incident["Service"]) &&
                           (incidentConfig.contactType === "RANDOM" || incidentConfig.contactType === incident["Contact type"]) &&
                           (incidentConfig.ftf === "RANDOM" || incidentConfig.ftf === incident["First time fix"]);
                });
                
                const selectedFromThisConfig = potentialIncidents.slice(0, incidentsForThisConfig);
                
                selectedFromThisConfig.forEach(incident => {
                    processedTaskNumbers.add(incident["Task Number"]);
                    selectedIncidents[sfMember][agent].push(incident);
                });
                
                incidentsForThisConfig = incidentsForThisConfig + (incidentsForThisConfig - selectedFromThisConfig.length);
            });

            const remainingIncidents = maxIncidents - selectedIncidents[sfMember][agent].length;
            if (remainingIncidents > 0) {
                const potentialIncidents = originalXlData.filter(incident => !processedTaskNumbers.has(incident["Task Number"]));
                const selectedRandomly = potentialIncidents.slice(0, remainingIncidents);
                selectedRandomly.forEach(incident => {
                    processedTaskNumbers.add(incident["Task Number"]);
                    selectedIncidents[sfMember][agent].push(incident);
                });
            }
        });
    }

    const rows = [];
    let previousSFMember = "";
    let previousAgent = "";

    for (const sfMember in selectedIncidents) {
        for (const agent in selectedIncidents[sfMember]) {
            for (let i = 0; i < incidentConfigs.length && selectedIncidents[sfMember][agent].length < maxIncidents; i++) {
                const incidentConfig = incidentConfigs[i];
                const filteredIncidents = originalXlData.filter((incident) => {
                    return (incidentConfig.service === "RANDOM" || incidentConfig.service === incident["Service"]) &&
                        (incidentConfig.contactType === "RANDOM" || incidentConfig.contactType === incident["Contact type"]) &&
                        (incidentConfig.ftf === "RANDOM" || incidentConfig.ftf === incident["First time fix"]) &&
                        !processedTaskNumbers.has(incident["Task Number"]);
                });
                filteredIncidents.forEach((incident) => {
                    if (selectedIncidents[sfMember][agent].length < maxIncidents) {
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