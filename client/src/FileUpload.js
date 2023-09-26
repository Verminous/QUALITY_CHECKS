import React, { useState, useMemo } from 'react';

const FileUpload = ({ onFileSelect, onConfigSubmit }) => {
    const fileInput = React.createRef();

    const services = useMemo(() => [
        'EMEIA Workplace',
        'Secure Internet Gateway (Global SIG)',
        'Identity and Access Management',
        'M365 Teams',
        'M365 Email',
        'M365 Apps',
        'Software Distribution (SCCM)',
        'Ask IT',
        'EMEIA Messaging',
        'Mobile Phones UK',
        'ZinZai Connect',
        'ForcePoint',
        'Network Service (CE/WEMEIA)',
        'M365 Sharepoint',
        'RANDOM'
    ], []);

    const contactTypes = useMemo(() => [
        'Self-service',
        'Phone - Unknown User',
        'Phone',
        'Chat',
        'RANDOM'
    ], []);

    // Default Configuration
    const defaultConfig = {
        incidentsPerAgent: 10,
        incidentConfigs: [
            { service: 'EMEIA Workplace', contactType: 'Phone', ftf: 'FALSE' },
            { service: 'EMEIA Workplace', contactType: 'Self-service', ftf: 'FALSE' },
            { service: 'Secure Internet Gateway (Global SIG)', contactType: 'Phone', ftf: 'FALSE' },
            { service: 'Secure Internet Gateway (Global SIG)', contactType: 'Self-service', ftf: 'FALSE' },
            { service: 'Identity and Access Management', contactType: 'Phone', ftf: 'FALSE' },
            { service: 'Identity and Access Management', contactType: 'Self-service', ftf: 'FALSE' },
            { service: 'RANDOM', contactType: 'Chat', ftf: 'FALSE' },
            { service: 'RANDOM', contactType: 'Phone', ftf: 'FALSE' },
            { service: 'RANDOM', contactType: 'Self-service', ftf: 'FALSE' },
            { service: 'RANDOM', contactType: 'Self-service', ftf: 'FALSE' }
        ]
    };

    const [config, setConfig] = useState(defaultConfig);

    const handleIncidentChange = (index, field, value) => {
        const updatedConfigs = [...config.incidentConfigs];
        updatedConfigs[index][field] = value;
        setConfig(prevConfig => ({ ...prevConfig, incidentConfigs: updatedConfigs }));
    };

    const handleFileInput = e => {
        // Get the selected file
        const file = fileInput.current.files[0];
        if (file) {
            onFileSelect(file);
        }
    };

    const handleConfigSubmit = () => {
        const sfMembersArray = sfMembers.split('\n').map(name => name.trim()).filter(name => name.length > 0);
        onConfigSubmit({ ...config, sfMembers: sfMembersArray });
    };
    
    const [sfMembers, setSfMembers] = useState(
        "Kempa, Martin\nSocha, Michał\nKrasowicz, Barbara\nSzczypior, Dawid\nSiemieniuk, Roman\nKalbarczyk, Jan\nLubonski, Piotr\nKucinska, Diana\nZiółkowski, Konrad\nKoplin, Krzysztof"
    );

    return (
        <div>
            <h2>Upload raw Excel file from Snow report</h2>
            <input type="file" ref={fileInput} onChange={handleFileInput} />
            <br /><br />

            <h2>Configuration</h2>
            <label>
                Incidents per Agent:
                <select name="incidentsPerAgent" value={config.incidentsPerAgent} onChange={(e) => setConfig({ ...config, incidentsPerAgent: parseInt(e.target.value) })}>
                    {[5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15].map(num => <option key={num} value={num}>{num}</option>)}
                </select>
            </label>
            <br /><br />

            <div className="config-grid">
                <strong>#</strong>
                <strong>Service:</strong>
                <strong>Type of contact:</strong>
                <strong>FTF:</strong>
                {config.incidentConfigs.map((incidentConfig, index) => (
                    <React.Fragment key={index}>
                        <span>{index + 1}</span>
                        <select value={incidentConfig.service} onChange={(e) => handleIncidentChange(index, 'service', e.target.value)}>
                            {services.map(service => <option key={service} value={service}>{service}</option>)}
                        </select>
                        <select value={incidentConfig.contactType} onChange={(e) => handleIncidentChange(index, 'contactType', e.target.value)}>
                            {contactTypes.map(type => <option key={type} value={type}>{type}</option>)}
                        </select>
                        <select value={incidentConfig.ftf} onChange={(e) => handleIncidentChange(index, 'ftf', e.target.value)}>
                            <option value="TRUE">TRUE</option>
                            <option value="FALSE">FALSE</option>
                        </select>
                    </React.Fragment>
                ))}
            </div>
            <br />

            <h2>Supporting Functions Team Members</h2>
            <textarea
                rows="10"
                cols="30"
                value={sfMembers}
                onChange={e => setSfMembers(e.target.value)}
                placeholder="Paste the list of SF team members here, one per line."
            />
            <br /><br />


            <button onClick={handleConfigSubmit}>Process</button>
            <br /><br />



            <h2>Download processed Excel file</h2>
            <button>Download</button>
            <br /><br /><br /><br />
        </div>
    );
};

export default FileUpload;
