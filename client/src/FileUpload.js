import React, { useState, useEffect, useMemo } from 'react';

const serverPort = process.env.REACT_APP_SERVER_PORT;
const currentHost = window.location.hostname;
const uploadUrl = `http://${currentHost}:${serverPort}/upload`;
const processUrl = `http://${currentHost}:${serverPort}/process`; // Adjust this if needed


const FileUpload = ({ onFileSelect, onConfigSubmit }) => {
    const fileInput = React.createRef();
 
    const services = useMemo(() => [
        'EMEIA Workplace',
        'Secure Internet Gateway (Global SIG)',
        'Identity and Access Management',
        'Identity Access Management (Finland)',
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

    const defaultConfig = {
        incidentsPerAgent: 10,
        incidentConfigs: [
            { service: 'EMEIA Workplace', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'EMEIA Workplace', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'Secure Internet Gateway (Global SIG)', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'Secure Internet Gateway (Global SIG)', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'Identity and Access Management', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'Identity and Access Management', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'RANDOM', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'RANDOM', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'RANDOM', contactType: 'RANDOM', ftf: 'RANDOM' },
            { service: 'RANDOM', contactType: 'RANDOM', ftf: 'RANDOM' }
        ]
    };

    const [config, setConfig] = useState(defaultConfig);

    useEffect(() => {
        setConfig(prevConfig => ({
            ...prevConfig,
            incidentConfigs: Array.from({ length: config.incidentsPerAgent }, (_, index) => (
                prevConfig.incidentConfigs[index] || { service: services[0], contactType: contactTypes[0], ftf: 'TRUE' }
            ))
        }));
    }, [config.incidentsPerAgent, contactTypes, services]);

    const handleIncidentChange = (index, field, value) => {
        const updatedConfigs = [...config.incidentConfigs];
        updatedConfigs[index][field] = value;
        setConfig(prevConfig => ({ ...prevConfig, incidentConfigs: updatedConfigs }));
    };

    const handleFileInput = async (e) => {
        const file = fileInput.current.files[0];
        if (file) {
            onFileSelect(file);
            const formData = new FormData();
            formData.append('file', file); 
            try {
                const response = await fetch(uploadUrl, {
                    method: 'POST',
                    body: formData,
                });
                const data = await response.json();
                setAgentAccounts(data.agentNames.join('\n'));
            } catch (error) {
                console.error('Error uploading file:', error);
            }
        }
    };

    const handleConfigSubmit = async () => {
        const sfMembersArray = sfMembers.split('\n').map(name => name.trim()).filter(name => name.length > 0);

        console.log('Submitting Config:', {
            ...config,
            sfMembers: sfMembersArray
          });

        try {
            const response = await fetch(processUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ ...config, sfMembers: sfMembersArray }),
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = process.env.REACT_APP_CLI_FILENAME;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
            } else {
                console.error('Failed to process the file.');
            }
        } catch (error) {
            console.error('Error processing file:', error);
        }
    };

    const [sfMembers, setSfMembers] = useState(
        "Kempa, Martin\nSocha, Michał\nKrasowicz, Barbara\nSzczypior, Dawid\nSiemieniuk, Roman\nKalbarczyk, Jan\nLubonski, Piotr\nKucinska, Diana\nZiółkowski, Konrad\nKoplin, Krzysztof"
    );
    const [agentAccounts, setAgentAccounts] = useState('');

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
                <strong>Contact Type:</strong>
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
                            <option value="RANDOM">RANDOM</option>
                        </select>
                    </React.Fragment>
                ))}
            </div>
            <br />
            <div className='all-accounts'>
                <div className='sf-accounts'>
                    <h2>SF team members</h2>
                    (Update/change manually)<br /><br />
                    <textarea
                        rows="10"
                        cols="30"
                        value={sfMembers}
                        onChange={e => setSfMembers(e.target.value)}
                        placeholder="Paste the list of SF team members here, one per line."
                    />
                </div>
                <div className='agent-accounts'>
                    <h2>Agent Accounts</h2>
                    ( Populated from raw Excel file<br /> - delete unwanted accounts/lines )<br /><br />
                    <textarea
                        rows="10"
                        value={agentAccounts}
                        onChange={e => setAgentAccounts(e.target.value)}
                        placeholder="Agent accounts will be automatically populated here after file upload."
                    />
                </div>
            </div>

            <h2>Process and Download Excel file</h2>
            <button onClick={handleConfigSubmit}>Process</button>
            <br /><br />

        </div>
    );
};

export default FileUpload;
