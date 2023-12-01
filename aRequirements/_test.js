const selectIncidentsByConfiguration = (
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
            const incidentConfig = incidentConfigs[i % incidentConfigs.length];
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
};
