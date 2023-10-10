const express = require("express"), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(bodyParser.json());
app.use(({ headers: { origin } }, res, next) => { const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`]; allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization')); next(); });
app.get("/", (req, res) => res.send("Server running!"));
app.listen(port, hostname, () => { console.log(`Server running at http://${hostname}:${port}/`); });

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => { lastUploadedFilePath = p; const workbook = xlsx.readFile(p), sheet = workbook.Sheets[workbook.SheetNames[0]], json = xlsx.utils.sheet_to_json(sheet); res.json({ agentNames: [...new Set(json.map(data => data["Taken By"]))] }); });
app.post("/process", upload.single("file"), async ({ body: config }, res) => { try { const workbook = xlsx.readFile(lastUploadedFilePath), sheet = workbook.Sheets[workbook.SheetNames[0]], originalXlData = xlsx.utils.sheet_to_json(sheet), { incidentConfigs, sfMembers, incidentsPerAgent } = config, selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, incidentsPerAgent, mapSFMembersToIncidentAgents(sfMembers, mapIncidentsByAgent(originalXlData))), rows = formatRowsForDownload(selectedIncidents); if (rows.length < incidentsPerAgent) { throw new Error("Not enough incidents matched the provided configuration"); } else { downloadFile(res, createAndWriteWorksheet(workbook, rows)); }; console.log(config); } catch (error) { console.error("Error in /process:", error, "Request body:", config, lastUploadedFilePath && "Last uploaded file path:", lastUploadedFilePath); res.status(500).send("Internal Server Error"); } });

const {
  selectIncidentsByConfiguration = async (
    originalXlData,
    incidentConfigs,
    maxIncidents,
    sfAgentMapping
  ) => {
    const selectedIncidents = {};
    Object.keys(sfAgentMapping).forEach((sfMember) => {
      selectedIncidents[sfMember] = {};
      const alreadySelected = {};
      sfAgentMapping[sfMember].forEach((agent) => {
        selectedIncidents[sfMember][agent] = [];
        alreadySelected[agent] = new Set();
        Array(maxIncidents)
          .fill()
          .forEach((_, i) => {
            const incidentConfig = incidentConfigs[i % incidentConfigs.length],
              potentialIncidents = [
                "Service",
                "Contact type",
                "First time fix",
              ].reduce(
                (incidents, field) =>
                  filterIncidentsByCriterion(
                    incidents,
                    field,
                    incidentConfig[field.toLowerCase()],
                    agent,
                    alreadySelected[agent]
                  ),
                [...originalXlData]
              ),
              selectedIncident = selectUniqueIncidentForAgent(
                potentialIncidents,
                alreadySelected[agent]
              );
            selectedIncident &&
              (selectedIncidents[sfMember][agent].push(selectedIncident),
              alreadySelected[agent].add(selectedIncident));
          });
      });
    });
    return selectedIncidents;
  },
  mapIncidentsByAgent = (data) =>
    data.reduce((acc, incident) => {
      const agent = incident["Taken By"];
      acc[agent] = acc[agent] || [];
      acc[agent].push(incident);
      return acc;
    }, {}),
  mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => {
    const sfAgentMapping = {},
      agents = Object.keys(incidentsByAgent);
    fisherYatesShuffle(agents).forEach((agent, index) => {
      const sfMember = sfMembers[index % sfMembers.length];
      sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent];
    });
    return sfAgentMapping;
  },
  filterByCriterion = (incidents, field, value, agent, alreadySelected) => {
    incidents = fisherYatesShuffle(incidents);
    value =
      value === "RANDOM"
        ? getRandomValue(incidents, field, alreadySelected)
        : value;
    const filtered = incidents.filter(
      (incident) =>
        !alreadySelected.has(incident) &&
        incident[field] === value &&
        incident["Taken By"] === agent
    );
    return filtered.length
      ? filtered
      : incidents.filter((incident) => incident["Taken By"] === agent);
  },
  selectUniqueIncident = (filteredIncidents, alreadySelected) => {
    const uniqueIncidents = filteredIncidents.filter(
      (incident) => !alreadySelected.has(incident)
    );
    return uniqueIncidents.length
      ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)]
      : null;
  },
  filterIncidentsByCriterion = (
    incidents,
    field,
    value,
    agent,
    alreadySelected,
    previousIncidents
  ) => {
    const filtered = filterByCriterion(
        incidents,
        field,
        value,
        agent,
        alreadySelected
      ),
      selectedIncident = selectUniqueIncident(filtered, alreadySelected);
    return selectedIncident
      ? [selectedIncident]
      : previousIncidents && previousIncidents.length > 0
      ? [selectUniqueIncident(previousIncidents, alreadySelected)]
      : [selectUniqueIncident(incidents, alreadySelected)];
  },
  getRandomValue = (incidents, field, alreadySelected) => {
    const uniqueValues = [...new Set(incidents.map((i) => i[field]))],
     unselectedValues = uniqueValues.filter(
      (value) => !alreadySelected.has(value)
    );
    return unselectedValues.length === 0
      ? uniqueValues[Math.floor(Math.random() * uniqueValues.length)]
      : unselectedValues[Math.floor(Math.random() * unselectedValues.length)];
  },
  fisherYatesShuffle = (array) => {
    array.forEach((_, i) => {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    });
    return array;
  },
  selectUniqueIncidentForAgent = (filteredIncidents, alreadySelected) => {
    const uniqueIncidents = filteredIncidents.filter(
      (incident) => !alreadySelected.has(incident)
    );
    return uniqueIncidents.length
      ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)]
      : null;
  },
  createAndWriteWorksheet = (workbook, rows) => {
    const newWorksheet = xlsx.utils.json_to_sheet(rows);
    workbook.Sheets["Processed List"]
      ? (workbook.Sheets["Processed List"] = newWorksheet)
      : xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List");
    const newFilePath = path.join(
      __dirname,
      "uploads",
      process.env.SERV_FILENAME
    );
    xlsx.writeFile(workbook, newFilePath);
    return newFilePath;
  },
  formatRowsForDownload = (selectedIncidents) => {
    let [previousSFMember, previousAgent] = [""];
    return Object.keys(selectedIncidents).flatMap((sfMember) =>
      Object.keys(selectedIncidents[sfMember]).flatMap((agent) =>
        selectedIncidents[sfMember][agent].map((incident) => {
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
        })
      )
    );
  },
  downloadFile = (res, newFilePath) => {
    res.download(newFilePath, filename, (err) => {
      if (err) throw new Error("Error sending the file: " + err);
    });
  },
} = {};
