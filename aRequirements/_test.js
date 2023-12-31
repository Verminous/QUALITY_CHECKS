const express = require("express"),
  ExcelJS = require("exceljs"),
  excel4node = require("excel4node"),
  cors = require("cors"),
  fs = require("fs"),
  multer = require("multer"),
  xlsx = require("xlsx"),
  bodyParser = require("body-parser"),
  path = require("path"),
  app = express(),
  upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const {
  UI_PORT: uiPort,
  SERV_PORT: port,
  HOSTNAME: hostname,
  SERV_FILENAME: filename,
} = process.env;

app.use(cors());
app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => {
  const allowed = [
    `http://localhost:${uiPort}`,
    `http://${hostname}:${uiPort}`,
  ];
  allowed.includes(origin) &&
    res.setHeader("Access-Control-Allow-Origin", origin);
  ["Methods", "Headers", "Credentials"].forEach((h) =>
    res.header(
      `Access-Control-Allow-${h}`,
      h == "Credentials"
        ? true
        : h == "Methods"
        ? "GET, POST, PUT, DELETE"
        : "Content-Type, Authorization"
    )
  );
  next();
});
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});

/* LOGGER */

const logDir = path.join(__dirname, "logs");
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir);
}
function logToFile(message) {
  console.log(message);
  fs.appendFile(path.join(logDir, "log.txt"), message + "\n", (err) => {
    if (err) throw err;
  });
}

/* ENDPOINTS */

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => {
  lastUploadedFilePath = p;
  const workbook = xlsx.readFile(p),
    sheet = workbook.Sheets[workbook.SheetNames[0]],
    json = xlsx.utils.sheet_to_json(sheet);
  res.json({ agentNames: [...new Set(json.map((data) => data["Taken By"]))] });
});
app.post("/process", upload.single("file"), async ({ body: config }, res) => {
  console.log(config);
  try {
    const workbook = xlsx.readFile(lastUploadedFilePath),
      sheet = workbook.Sheets[workbook.SheetNames[0]],
      originalXlData = xlsx.utils.sheet_to_json(sheet),
      {
        incidentsPerAgent,
        incidentConfigs,
        sfMembers,
        agentNames,
        randomServices,
      } = config,
      selectedIncidents = await selectIncidentsByConfiguration(
        originalXlData,
        incidentConfigs,
        incidentsPerAgent,
        mapSFMembersToIncidentAgents(
          sfMembers,
          mapIncidentsByAgent(originalXlData)
        ),
        randomServices, config
      ),
      rows = formatRowsForDownload(selectedIncidents);
    if (rows.length < incidentsPerAgent) {
      throw new Error(
        "Not enough incidents matched the provided configuration"
      );
    } else {
      downloadFile(res, createAndWriteWorksheet(workbook, rows));
    }
    console.log(config);
  } catch (error) {
    console.error(
      "Error in /process:",
      error,
      "Request body:",
      config,
      lastUploadedFilePath && "Last uploaded file path:",
      lastUploadedFilePath
    );
    res.status(500).send("Internal Server Error");
  }
});

/* DATA PROCESSING */

const {
  triedServices = {},

  getRandomValue = (incidents, field, alreadySelected) => {
    const uniqueValues = [...new Set(incidents.map((i) => i[field]))];
    const unselectedValues = uniqueValues.filter(
      (value) => !alreadySelected.has(value)
    );
    const valuesToUse =
      unselectedValues.length > 0 ? unselectedValues : uniqueValues;
    return valuesToUse[Math.floor(Math.random() * valuesToUse.length)];
  },
  fisherYatesShuffle = (array) => {
    array.forEach((_, i) => {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    });
    return array;
  },
  mapIncidentsByAgent = (data) =>
    data.reduce((acc, incident) => {
      const agent = incident["Taken By"];
      acc[agent] = acc[agent] || [];
      const existingIncidentIndex = acc[agent].findIndex(
        (i) => i["Task Number"] === incident["Task Number"]
      );
      if (existingIncidentIndex !== -1) {
        acc[agent][existingIncidentIndex] = {
          ...acc[agent][existingIncidentIndex],
          ...incident,
        };
      } else {
        acc[agent].push(incident);
      }
      return acc;
    }, {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => {
    const sfAgentMapping = {},
      agents = config.agentNames;
    fisherYatesShuffle(agents).forEach((agent, index) => {
      const sfMember = sfMembers[index % sfMembers.length];
      sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent];
    });
    return sfAgentMapping;
  },
  selectUniqueIncidentForAgent = (
    filteredIncidents = [],
    alreadySelected,
    originalIncidents = []
  ) => {
    let uniqueIncidents = filteredIncidents.filter(
      (incident) => !alreadySelected.has(incident["Task Number"])
    );
    if (!uniqueIncidents.length) {
      uniqueIncidents = originalIncidents.filter(
        (incident) => !alreadySelected.has(incident["Task Number"])
      );
    }
    let selectedIncident = uniqueIncidents.length
      ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)]
      : null;
    if (selectedIncident) {
      alreadySelected.add(selectedIncident["Task Number"]);
    } else {
      logToFile(
        "All original incidents have been selected for the current agent."
      );
    }
    return selectedIncident;
  },

  filterByCriterion = (
    incidents,
    field,
    value,
    agent,
    alreadySelected,
    triedValues = new Set(),
    randomServices = []
  ) => {
    logToFile(
      `filterByCriterion - Start: field=${field}, value=${value}, agent=${agent}`
    );
    if (field === "Service" && value === "RANDOM") {
      const untriedServices = randomServices.filter(
        (service) => !triedValues.has(service)
      );
      logToFile(`filterByCriterion - Untried services: ${untriedServices}`);
      if (untriedServices.length > 0) {
        value =
          untriedServices[Math.floor(Math.random() * untriedServices.length)];
      } else {
        value = getRandomValue(incidents, field, alreadySelected);
        alreadySelected.add(value);
      }
      logToFile(`filterByCriterion - New value for RANDOM service: ${value}`);
    } else if (value === "RANDOM") {
      value = getRandomValue(incidents, field, alreadySelected);
      alreadySelected.add(value);
    }
    triedValues.add(value);
    triedServices[agent] = triedValues;
    const filtered = incidents.filter((incident) => {
      const matches =
        !alreadySelected.has(incident) &&
        incident[field] === value &&
        incident["Taken By"] === agent;
      return matches;
    });
    logToFile(`filterByCriterion - Filtered incidents: ${filtered.length}`);
    if (!filtered.length) {
      const allValues = incidents.map((i) => i[field]);
      const untriedValues = allValues.filter(
        (value) => !triedValues.has(value)
      );
      if (untriedValues.length === 0) {
        return [];
      }
      const untriedServices = randomServices.filter(
        (service) => !triedValues.has(service)
      );
      value =
        field === "Service" && untriedServices.length > 0
          ? untriedServices[Math.floor(Math.random() * untriedServices.length)]
          : getRandomValue(incidents, field, alreadySelected);
      if (!triedValues.has(value)) {
        return filterByCriterion(
          incidents,
          field,
          value,
          agent,
          alreadySelected,
          triedValues,
          randomServices
        );
      }
    }
    return filtered;
  },
  filterIncidentsByFields = (
    incidents,
    fieldCriteria,
    agent,
    alreadySelected,
    randomServices
  ) => {
    logToFile(
      `filterIncidentsByFields - Start: fieldCriteria=${JSON.stringify(
        fieldCriteria
      )}, agent=${agent}, randomServices=${JSON.stringify(randomServices)}`
    );
    return Object.entries(fieldCriteria).reduce(
      (currentIncidents, [field, value]) => {
        logToFile(
          `filterIncidentsByFields - Processing field: ${field}, value: ${value}`
        );
        return filterByCriterion(
          currentIncidents,
          field,
          value,
          agent,
          alreadySelected,
          triedServices[agent] || new Set(),
          randomServices
        );
      },
      incidents
    );
  },
  selectIncidentsByConfiguration = (
    originalXlData,
    incidentConfigs,
    maxIncidents,
    sfAgentMapping,
    randomServices
  ) => {
    (!Array.isArray(originalXlData) || !originalXlData.length) &&
      (() => {
        throw new Error("Invalid originalXlData");
      })();
    const fieldToConfigKey = {
      Service: "service",
      "Contact type": "contactType",
      "First time fix": "ftf",
    };
    return Object.entries(sfAgentMapping).reduce(
      (selectedIncidents, [sfMember, agents]) => {
        selectedIncidents[sfMember] = agents.reduce((agentIncidents, agent) => {
          const alreadySelected = new Set();
          agentIncidents[agent] = Array(maxIncidents)
            .fill()
            .reduce((incidents, _, i) => {
              const incidentConfig =
                incidentConfigs[i % incidentConfigs.length];
              let fieldCriteria = {};
              ["Service", "Contact type", "First time fix"].forEach((field) => {
                const configKey = fieldToConfigKey[field];
                if (incidentConfig[configKey]) {
                  fieldCriteria[field] = incidentConfig[configKey];
                }
              });
              const potentialIncidents = filterIncidentsByFields(
                originalXlData,
                fieldCriteria,
                agent,
                alreadySelected,
                randomServices
              );
              if (potentialIncidents.length) {
                const uniqueIncident = selectUniqueIncidentForAgent(
                  potentialIncidents,
                  alreadySelected
                );
                if (uniqueIncident) {
                  incidents.push(uniqueIncident);
                  alreadySelected.add(uniqueIncident["Task Number"]);
                } else {
                  logToFile(
                    `Warning: No incidents available for fallback for agent ${agent}.`
                  );
                }
              }
              logToFile(
                `selectIncidentsByConfiguration_2 - Selected ${incidents.length} incidents for agent ${agent}`
              );
              return incidents;
            }, []);
          return agentIncidents;
        }, {});
        return selectedIncidents;
      },
      {}
    );
  },
} = {};

/* WRITE + DOWNLOAD */
const {
  createAndWriteWorksheet = (workbook, rows) => {
    const { json_to_sheet, book_append_sheet } = xlsx.utils;
    const { join } = path;
    const { SERV_FILENAME } = process.env;
    const newWorksheet = json_to_sheet(rows);
    const sheetName = "Processed List";
    newWorksheet["!cols"] = [
      { wch: 3 },
      { wch: 20 },
      { wch: 25 },
      { wch: 15 },
      { wch: 30 } /*{ wch: 20 }, { wch: 10 },*/,
    ];
    if (!workbook.Sheets[sheetName]) {
      book_append_sheet(workbook, newWorksheet, sheetName);
    } else {
      workbook.Sheets[sheetName] = newWorksheet;
    }
    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = workbook.Sheets[firstSheetName];
    firstSheet["!cols"] = [
      { wch: 15 },
      { wch: 30 },
      { wch: 30 },
      { wch: 30 },
      { wch: 15 },
      { wch: 10 },
    ];
    const newFilePath = join(__dirname, "uploads", SERV_FILENAME);
    xlsx.writeFile(workbook, newFilePath);
    return newFilePath;
  },
  formatRowsForDownload = (selectedIncidents) => {
    let [previousSFMember, previousAgent] = ["", ""];
    let rows = [];
    let incidentCount = 1;
    Object.entries(selectedIncidents).forEach(([sfMember, agents], sfIndex) => {
      Object.entries(agents).forEach(([agent, incidents], agentIndex) => {
        incidents.forEach((incident) => {
          const row = {
            "n.":
              previousAgent === agent ? incidentCount++ : (incidentCount = 1),
            "SF Member": previousSFMember === sfMember ? "" : sfMember,
            Agent: previousAgent === agent ? "" : agent,
            ...[
              "Task Number",
              "Service" /* , "Contact type", "First time fix" */,
            ].reduce((acc, key) => ({ ...acc, [key]: incident[key] }), {}),
          };
          rows.push(row);
          if (previousAgent !== agent) {
            incidentCount = 2;
          }
          [previousSFMember, previousAgent] = [sfMember, agent];
        });
        agentIndex < Object.keys(agents).length - 1 ? rows.push({}) : null;
      });
      sfIndex < Object.keys(selectedIncidents).length - 1
        ? rows.push({}, {})
        : null;
    });
    rows.splice(0, 0, {});
    return rows;
  },
  downloadFile = (res, newFilePath) => {
    const errorHandler = (err, message) => {
      if (err) throw new Error(`${message} ${err}`);
    };
    const unlinkFile = (path, message) =>
      fs.unlink(path, (err) => errorHandler(err, message));
    res.download(newFilePath, filename, (err) => {
      errorHandler(err, "Error sending the file:");
      unlinkFile(newFilePath, "Error deleting the processed file:");
      unlinkFile(lastUploadedFilePath, "Error deleting the temporary file:");
    });
  },
} = {};
