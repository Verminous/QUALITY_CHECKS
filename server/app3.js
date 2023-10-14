const express = require("express"), fs = require('fs'), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => { const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`]; allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization')); next(); });
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => { console.log(`Server running at http://${hostname}:${port}/`); });

// Define the log directory path
const logDir = path.join(__dirname, 'logs');

// Check if the log directory exists, if not create it
if (!fs.existsSync(logDir)) {
    fs.mkdirSync(logDir);
}

function logToFile(message) {
    console.log(message);
    fs.appendFile(path.join(logDir, 'log.txt'), message + '\n', (err) => {
        if (err) throw err;
    });
}

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => { lastUploadedFilePath = p; logToFile('Reading Excel file...'); const workbook = xlsx.readFile(p), sheet = workbook.Sheets[workbook.SheetNames[0]], json = xlsx.utils.sheet_to_json(sheet); logToFile('Excel file read successfully.'); res.json({ agentNames: [...new Set(json.map(data => data["Taken By"]))] }); });
app.post("/process", upload.single("file"), async ({ body: config }, res) => {
    try {
        const workbook = xlsx.readFile(lastUploadedFilePath),
            sheet = workbook.Sheets[workbook.SheetNames[0]],
            originalXlData = xlsx.utils.sheet_to_json(sheet),
            { incidentConfigs, sfMembers, incidentsPerAgent } = config;

        logToFile(`Processing with config: ${JSON.stringify(config)}`);

        const selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, incidentsPerAgent, mapSFMembersToIncidentAgents(sfMembers, mapIncidentsByAgent(originalXlData))),
            rows = formatRowsForDownload(selectedIncidents);

        if (rows.length < incidentsPerAgent) {
            logToFile(`Error: Not enough incidents matched the provided configuration. Rows: ${rows.length}, Incidents per agent: ${incidentsPerAgent}`);
            throw new Error("Not enough incidents matched the provided configuration");
        } else {
            downloadFile(res, createAndWriteWorksheet(workbook, rows));
        }
    } catch (error) {
        logToFile(`Error in /process: ${error}`);
        res.status(500).send("Internal Server Error");
    }
});

/* PROCESSING DATA */

const {
    getRandomItem = (array) => {
        return array[Math.floor(Math.random() * array.length)];
    },
    fisherYatesShuffle = (array) => {
        array.forEach((_, i) => {
            const j = Math.floor(Math.random() * (i + 1));
            [array[i], array[j]] = [array[j], array[i]];
        });
        return array;
    },
    mapIncidentsByAgent = (data) => data.reduce((acc, incident) => {
        const agent = incident["Taken By"];
        acc[agent] = acc[agent] || [];
        acc[agent].push(incident);
        return acc;
    }, {}),
    mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => {
        const sfAgentMapping = {};
        const agents = Object.keys(incidentsByAgent);
        fisherYatesShuffle(agents).forEach((agent, index) => {
            const sfMember = sfMembers[index % sfMembers.length];
            sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent];
        });
        return sfAgentMapping;
    },
    hasIncidentBeenAssigned = (incident, alreadySelectedIncidents) => {
        return alreadySelectedIncidents.has(incident.incidentId);
    },

    filterIncidentsBy = (incidents, property, value) => {
        return incidents.filter(incident => incident[property] === value);
    },
    selectFromPool = (incidents, property, value, alreadySelected = new Set()) => {
        logToFile(`Starting selectFromPool with property: ${property}, value: ${value}`);

        const uniqueValues = [...new Set(incidents.map(incident => incident[property]))];
        let remainingPool = [...uniqueValues];
        let selectedValue;
        while (remainingPool.length) {
            selectedValue = value === "RANDOM" ? getRandomItem(remainingPool) : value;
            const filteredIncidents = filterIncidentsBy(incidents, property, selectedValue);
            if (filteredIncidents.length) return { value: selectedValue, incidents: filteredIncidents };
            remainingPool = remainingPool.filter(val => val !== selectedValue);
            if (value !== "RANDOM") break;
        }
        logToFile(`Finished selectFromPool. Selected value: ${selectedValue}, incidents: ${JSON.stringify(incidents)}`);
        return { value: selectedValue, incidents: filteredIncidents };
    },
    XYZM = (incidents, serviceValue = "RANDOM", contactTypeValue = "RANDOM", firstTimeFixValue = "RANDOM", alreadySelected = new Set()) => {
        logToFile(`Starting XYZM with serviceValue: ${serviceValue}, contactTypeValue: ${contactTypeValue}, firstTimeFixValue: ${firstTimeFixValue}`);
        const serviceResult = selectFromPool(incidents, 'Service', serviceValue);
        if (!serviceResult) {
            console.warn("No incidents available for the agent based on the Service selection");
            return null;
        }
        const contactTypeResult = selectFromPool(serviceResult.incidents, 'Contact type', contactTypeValue);
        if (!contactTypeResult) return XYZM(incidents, "RANDOM");
        const firstTimeFixResult = selectFromPool(contactTypeResult.incidents, 'First time fix', firstTimeFixValue);
        if (!firstTimeFixResult) return XYZM(incidents, serviceResult.value, "RANDOM");
        let uniqueIncidents = firstTimeFixResult.incidents.filter((incident) => !hasIncidentBeenAssigned(incident, alreadySelected));
        let selectedIncident = getRandomItem(uniqueIncidents);
        if (!uniqueIncidents.length) return null;
        alreadySelected.add(selectedIncident.incidentId);
        logToFile(`Finished XYZM. Selected incident: ${JSON.stringify(selectedIncident)}`);
        return selectedIncident;
    },
    selectIncidentsByConfiguration = (originalXlData, incidentConfigs, sfMembers) => {
        const selectedIncidents = {};
        const maxIterations = 100000; 

        sfMembers.forEach((sfMember) => {
            selectedIncidents[sfMember] = {};
            incidentConfigs.forEach((config) => {
                const agents = sfAgentMapping[config.SFMember];
                let iterations = 0;
                for (let i = 0; i < incidentsPerAgent; i++) {
                    while (iterations < maxIterations) {
                        const selectedIncident = XYZM(incidentsByAgent[getRandomItem(agents)], config.Service, config["Contact type"], config["First time fix"], alreadySelectedIncidents);
                        if (selectedIncident) {
                            if (!selectedIncidents[sfMember][selectedIncident.TakenBy]) {
                                selectedIncidents[sfMember][selectedIncident.TakenBy] = [];
                            }
                            selectedIncidents[sfMember][selectedIncident.TakenBy].push(selectedIncident);
                            break;
                        }
                        iterations++;
                    }
                    if (iterations === maxIterations) {
                        throw new Error("No matching incidents found after maximum iterations");
                    }
                }
            });
        });
        logToFile(`Finished incident selection. Selected incidents: ${JSON.stringify(output)}`);
        return selectedIncidents;
    },
} = {};

/* WRITE + DOWNLOAD */
const {
    createAndWriteWorksheet = (workbook, rows) => {
        if (Array.isArray(rows) && rows.every(row => typeof row === 'object')) {
            const newWorksheet = xlsx.utils.json_to_sheet(rows);
            if (workbook.Sheets["Processed List"]) {
                workbook.Sheets["Processed List"] = newWorksheet;
            } else {
                xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List");
            }
            if (typeof process.env.SERV_FILENAME === 'string') {
                const newFilePath = path.join(
                    __dirname,
                    "uploads",
                    process.env.SERV_FILENAME
                );
                xlsx.writeFile(workbook, newFilePath);
                return newFilePath;
            } else {
                console.error('Invalid SERV_FILENAME:', process.env.SERV_FILENAME);
                return 'AskIT - QCH_processed.xlsx';
            }
        } else {
            console.error('Invalid rows:', rows);
            return 'AskIT - QCH_processed.xlsx';
        }
    },
    formatRowsForDownload = (selectedIncidents) => {
        let [previousSFMember, previousAgent] = [""];

        return Object.keys(selectedIncidents).flatMap((sfMember) =>
            Object.keys(selectedIncidents[sfMember]).flatMap((agent) => {
                if (Array.isArray(selectedIncidents[sfMember][agent])) {
                    return selectedIncidents[sfMember][agent].map((incident) => {
                        if (!incident) {
                            logToFile(`Error: Incident is undefined for SFMember: ${sfMember}, Agent: ${agent}`);
                            return;
                        }

                        const row = {
                            "SF Member": previousSFMember === sfMember ? "" : sfMember,
                            Agent: previousAgent === agent ? "" : agent,
                            "Task Number": incident["Task Number"],
                            Service: incident["Service"],
                            "Contact type": incident["Contact type"],
                            "First time fix": incident["First time fix"],
                        };

                        [previousSFMember, previousAgent] = [sfMember, agent];
                        return row;
                    });
                } else {
                    logToFile(`Error: selectedIncidents[${sfMember}][${agent}] is not an array`);
                }
            })
        ).filter(row => row); 
    },
    downloadFile = (res, newFilePath) => {
        res.download(newFilePath, filename, (err) => {
            if (err) throw new Error("Error sending the file: " + err);
            fs.unlink(newFilePath, (err) => {
                if (err) throw new Error("Error deleting the processed file: " + err);
                console.log("Processed file deleted successfully");
            });
            fs.unlink(lastUploadedFilePath, (err) => {
                if (err) throw new Error("Error deleting the temporary file: " + err);
                console.log("Temporary file deleted successfully");
            });
        });
    },
} = {};