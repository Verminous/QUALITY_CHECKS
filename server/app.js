const express = require("express"), multer = require("multer"), xlsx = require("xlsx"), bodyParser = require("body-parser"), path = require("path"), app = express(), upload = multer({ dest: "uploads/" });
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { UI_PORT: uiPort, SERV_PORT: port, HOSTNAME: hostname, SERV_FILENAME: filename } = process.env;

app.use(bodyParser.json());

app.use(({headers: {origin}}, res, next) => {
  const allowed = [`http://localhost:${uiPort}`, `http://${hostname}:${uiPort}`];
  allowed.includes(origin) && res.setHeader('Access-Control-Allow-Origin', origin);
  ['Methods', 'Headers', 'Credentials'].forEach(h => res.header(`Access-Control-Allow-${h}`, h == 'Credentials' ? true : h == 'Methods' ? 'GET, POST, PUT, DELETE' : 'Content-Type, Authorization'));
  next();
});

app.get("/", (req, res) => res.send("Server running!"));

let lastUploadedFilePath;
app.post("/upload", upload.single("file"), ({ file: { path: p } }, res) => {
  lastUploadedFilePath = p;
  res.json({ agentNames: [...new Set(xlsx.utils.sheet_to_json(xlsx.readFile(p).Sheets[xlsx.readFile(p).SheetNames[0]]).map(data => data["Taken By"]))] });
});

app.post("/process", upload.single("file"), async ({ body: config }, res) => {
  try {
    const workbook = xlsx.readFile(lastUploadedFilePath),
      sheetName = workbook.SheetNames[0],
      originalXlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]),
      { incidentConfigs, sfMembers, incidentsPerAgent } = config,
      incidentsByAgent = mapIncidentsByAgent(originalXlData),
      sfAgentMapping = mapSFMembersToIncidentAgents(sfMembers, incidentsByAgent),
      selectedIncidents = await selectIncidentsByConfiguration(originalXlData, incidentConfigs, incidentsPerAgent, sfAgentMapping),
      rows = formatRowsForDownload(selectedIncidents);
    if (rows.length < incidentsPerAgent) throw new Error("Not enough incidents matched the provided configuration");
    const newFilePath = createAndWriteWorksheet(workbook, rows);
    downloadFile(res, newFilePath);
  } catch (error) {
    console.error("Error in /process:", error);
    console.error("Request body:", config);
    lastUploadedFilePath && console.error("Last uploaded file path:", lastUploadedFilePath);
    res.status(500).send("Internal Server Error");
  }
});

const selectIncidentsByConfiguration = async (originalXlData, incidentConfigs, maxIncidents, sfAgentMapping) => {
  const [selectedIncidents, alreadySelected] = [{}, {}];
  Object.keys(sfAgentMapping).forEach(sfMember => {
    selectedIncidents[sfMember] = {};
    sfAgentMapping[sfMember].forEach(agent => {
      selectedIncidents[sfMember][agent] = [];
      alreadySelected[agent] = new Set();
      Array(maxIncidents).fill().map((_, i) => {
        const incidentConfig = incidentConfigs[i % incidentConfigs.length],
          potentialIncidents = ['Service', 'Contact type', 'First time fix'].reduce((incidents, field) => filterIncidentsByCriterion(incidents, field, incidentConfig[field.toLowerCase()], agent, alreadySelected[agent]), [...originalXlData]),
          selectedIncident = selectUniqueIncidentForAgent(potentialIncidents, alreadySelected[agent]);
        selectedIncident && (selectedIncidents[sfMember][agent].push(selectedIncident), alreadySelected[agent].add(selectedIncident));
      });
    });
  });
  return selectedIncidents;
}

const mapIncidentsByAgent = data =>
  data.reduce((acc, incident) => {
    const agent = incident["Taken By"];
    acc[agent] = acc[agent] || [];
    acc[agent].push(incident);
    return acc;
  }, {});

const mapSFMembersToIncidentAgents = (sfMembers, incidentsByAgent) => {
  const sfAgentMapping = {};
  Object.keys(incidentsByAgent).sort(() => 0.5 - Math.random()).forEach((agent, index) => {
    const sfMember = sfMembers[index % sfMembers.length];
    sfAgentMapping[sfMember] = [...(sfAgentMapping[sfMember] || []), agent];
  });
  return sfAgentMapping;
}

const filterIncidentsByCriterion = (incidents, field, value, agent, alreadySelected) => {
  incidents = fisherYatesShuffle(incidents);
  value = (value === 'RANDOM') ? getRandomValue(incidents, field) : value;
  const filtered = incidents.filter(incident => !alreadySelected.has(incident) && incident[field] === value && incident['Taken By'] === agent);
  return filtered.length ? filtered : incidents.filter(incident => incident['Taken By'] === agent);
}

const getRandomValue = (incidents, field) => [...new Set(incidents.map(i => i[field]))][Math.floor(Math.random() * incidents.length)];

const fisherYatesShuffle = array => {
  array.forEach((i) => { const j = Math.floor(Math.random() * (i + 1));[array[i], array[j]] = [array[j], array[i]]; });
  return array;
};

const selectUniqueIncidentForAgent = (filteredIncidents, alreadySelected) => {
  const uniqueIncidents = filteredIncidents.filter(incident => !alreadySelected.has(incident));
  return uniqueIncidents.length ? uniqueIncidents[Math.floor(Math.random() * uniqueIncidents.length)] : null;
};

const createAndWriteWorksheet = (workbook, rows) => {
  const newWorksheet = xlsx.utils.json_to_sheet(rows);
  workbook.Sheets["Processed List"] ? workbook.Sheets["Processed List"] = newWorksheet : xlsx.utils.book_append_sheet(workbook, newWorksheet, "Processed List");
  const newFilePath = path.join(__dirname, "uploads", process.env.SERV_FILENAME);
  xlsx.writeFile(workbook, newFilePath);
  return newFilePath;
};

const formatRowsForDownload = selectedIncidents => {
  let [previousSFMember, previousAgent] = [""];
  return Object.keys(selectedIncidents).flatMap(sfMember =>
    Object.keys(selectedIncidents[sfMember]).flatMap(agent =>
      selectedIncidents[sfMember][agent].map(incident => {
        const row = {
          "SF Member": previousSFMember === sfMember ? "" : sfMember,
          Agent: previousAgent === agent ? "" : agent,
          "Task Number": incident["Task Number"],
          Service: incident["Service"],
          "Contact type": incident["Contact type"],
          "First time fix": incident["First time fix"],
        };
        [previousSFMember, previousAgent] = [previousSFMember !== sfMember ? sfMember : previousSFMember, previousAgent !== agent ? agent : previousAgent];
        return row;
      })
    )
  );
};

const downloadFile = (res, newFilePath) => {
  res.download(newFilePath, filename, (err) => {
    if (err) throw new Error("Error sending the file: " + err);
  });
};

app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});