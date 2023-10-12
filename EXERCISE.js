const incidents = [ /* ... sample data ... */ ];
const incidentConfigs = [ /* ... configuration from UI ... */ ];
const incidentsPerAgent = 10; /* example value */
const assignedIncidents = new Map();
const sfMembers = [ /* ... SF members data ... */ ];
const agents = [ /* ... agents data ... */ ];
const sfMemberAgents = new Map();
const sfMemberIncidents = new Map();

function getRandomItem(array) {
    return array[Math.floor(Math.random() * array.length)];
}

function filterIncidentsBy(incidents, property, value) {
    return incidents.filter(incident => incident[property] === value);
}

function selectFromPool(incidents, property, value) {
    const uniqueValues = [...new Set(incidents.map(incident => incident[property]))];
    let remainingPool = [...uniqueValues];
    let selectedValue;

    while (remainingPool.length) {
        selectedValue = value === "RANDOM" ? getRandomItem(remainingPool) : value;
        const filteredIncidents = filterIncidentsBy(incidents, property, selectedValue);
        
        if (filteredIncidents.length) 
            return { value: selectedValue, incidents: filteredIncidents };
        
        remainingPool = remainingPool.filter(val => val !== selectedValue);
        if (value !== "RANDOM") break; 
    }

    return null;
}

function hasIncidentBeenAssigned(incident) {
    const agent = incident['Taken By']; 
    const taskNumber = incident['Task Number'];
    if (!assignedIncidents.has(agent)) return false;
    return assignedIncidents.get(agent).has(taskNumber);
}

function assignIncidentToAgent(incident) {
    const agent = incident['Taken By'];
    const taskNumber = incident['Task Number'];
    if (!assignedIncidents.has(agent)) assignedIncidents.set(agent, new Set());
    assignedIncidents.get(agent).add(taskNumber);
}

function assignAgentsToSFMembers() {
    const sfMemberCount = sfMembers.length;
    let remainingAgents = [...agents];

    for (const sfMember of sfMembers) {
        sfMemberAgents.set(sfMember, []);
        sfMemberIncidents.set(sfMember, []); 
    }

    let sfMemberIndex = 0;
    while (remainingAgents.length) {
        const currentSFMember = sfMembers[sfMemberIndex];
        const currentAgent = remainingAgents.pop();
        sfMemberAgents.get(currentSFMember).push(currentAgent);
        sfMemberIndex = (sfMemberIndex + 1) % sfMemberCount;
    }
}

assignAgentsToSFMembers();

function XYZM(incidents, serviceValue = "RANDOM", contactTypeValue = "RANDOM", firstTimeFixValue = "RANDOM") {
    const serviceResult = selectFromPool(incidents, 'Service', serviceValue);
    if (!serviceResult) {
        console.warn("No incidents available for the agent based on the Service selection");
        return null;
    }
    
    const contactTypeResult = selectFromPool(serviceResult.incidents, 'Contact type', contactTypeValue);
    if (!contactTypeResult) return XYZM(incidents, "RANDOM"); 

    const firstTimeFixResult = selectFromPool(contactTypeResult.incidents, 'First time fix', firstTimeFixValue);
    if (!firstTimeFixResult) return XYZM(incidents, serviceResult.value, "RANDOM"); 
    
    const uniqueIncidents = firstTimeFixResult.incidents.filter(incident => !hasIncidentBeenAssigned(incident));
    if (!uniqueIncidents.length) return null; 

    const selectedIncident = getRandomItem(uniqueIncidents);
    assignIncidentToAgent(selectedIncident);

    return selectedIncident;
}

assignAgentsToSFMembers();

for (let i = 0; i < incidentsPerAgent; i++) {
    for (const config of incidentConfigs) {
        for (const agent of agents) {
            const agentIncidents = incidents.filter(incident => incident['Taken By'] === agent);
            const selectedIncident = XYZM(agentIncidents, config.service, config.contactType, config.ftf);
            if (!selectedIncident) {
                console.warn(`No more incidents available for agent ${agent}`);
                break;
            }

            for (const [sfMember, sfMemberAgentList] of sfMemberAgents) {
                if (sfMemberAgentList.includes(agent)) {
                    sfMemberIncidents.get(sfMember).push(selectedIncident);
                    break;
                }
            }
        }
    }
}

console.log(assignedIncidents);
console.log(sfMemberIncidents);