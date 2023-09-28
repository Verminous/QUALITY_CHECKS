const express = require('express'); 
const multer = require('multer'); 
const app = express(); 
const port = 3001; 
const xlsx = require('xlsx'); 
const upload = multer({ dest: 'uploads/' }); 
const bodyParser = require('body-parser'); 
const cors = require('cors'); 
const fs = require('fs'); 
const path = require('path');

app.use(bodyParser.json()); 
app.use(cors()); 

app.get('/', (req, res) => { res.send('Server running!'); }); 
let lastUploadedFilePath; 

app.post('/upload', upload.single('file'), (req, res) => { 
    lastUploadedFilePath = req.file.path; 
    const workbook = xlsx.readFile(lastUploadedFilePath); 
    const sheet_name_list = workbook.SheetNames; 
    const xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); 
    const agentNames = [...new Set(xlData.map(data => data['Taken By']))]; 
    res.json({ agentNames }); 
});

app.post('/process', upload.single('file'), async (req, res) => { 
    const config = req.body; 
    const workbook = xlsx.readFile(lastUploadedFilePath); 
    const sheetName = workbook.SheetNames[0]; 
    const xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]); 

    const incidentsByAgent = {}; 
    xlData.forEach(incident => { 
        const agent = incident['Taken By']; 
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

    try { 
        for (const sfMember in sfAgentMapping) { 
            selectedIncidents[sfMember] = {}; 
            sfAgentMapping[sfMember].forEach(agent => { 
                selectedIncidents[sfMember][agent] = []; 
                const agentIncidents = incidentsByAgent[agent]; 
                config.incidentConfigs.forEach(incidentConfig => { 
                    const matchedIncidents = agentIncidents.filter(incident => 
                        incident.Service === incidentConfig.service && 
                        incident['Contact type'] === incidentConfig.contactType && 
                        incident['First time fix'] === incidentConfig.ftf
                    );
                    for (let i = 0; i < matchedIncidents.length && selectedIncidents[sfMember][agent].length < config.incidentsPerAgent; i++) { 
                        selectedIncidents[sfMember][agent].push(matchedIncidents[i]); 
                    } 
                }); 
            }); 
        } 
    } catch (error) { 
        console.error('Error during data processing:', error); 
    } 

    const rows = []; 
    const headers = { 
        'SF Member': 'SF Member', 
        'Agent': 'Agent', 
        'Task Number': 'Task Number', 
        'Service': 'Service', 
        'Contact Type': 'Contact Type', 
        'First Time Fix': 'First Time Fix' 
    };
    rows.push(headers);

    for (const sfMember in selectedIncidents) { 
        for (const agent in selectedIncidents[sfMember]) { 
            selectedIncidents[sfMember][agent].forEach(incident => { 
                rows.push({ 
                    'SF Member': sfMember, 
                    'Agent': agent, 
                    'Task Number': incident['Task Number'], 
                    'Service': incident['Service'], 
                    'Contact Type': incident['Contact type'], 
                    'First Time Fix': incident['First time fix'] 
                }); 
            }); 
        } 
    } 

    const newWorkbook = workbook; 
    const newWorksheet = xlsx.utils.json_to_sheet(rows); 

    if (!newWorkbook.Sheets["Processed List"]) { 
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Processed List"); 
    } else { 
        newWorkbook.Sheets["Processed List"] = newWorksheet; 
    } 

    const newFilePath = path.join(__dirname, 'uploads', 'AskIT - QCH_processed.xlsx'); 
    xlsx.writeFile(newWorkbook, newFilePath); 
    res.download(newFilePath, 'AskIT - QCH_processed.xlsx'); 
});

app.listen(port, () => { 
    console.log(`Server running at http://localhost:${port}/`); 
});
