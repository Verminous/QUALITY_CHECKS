import React, { useState, useEffect, useMemo } from 'react';

const serverPort = process.env.REACT_APP_SERVER_PORT;
const currentHost = window.location.hostname;
const uploadUrl = `http://${currentHost}:${serverPort}/upload`;
const processUrl = `http://${currentHost}:${serverPort}/process`;

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

    const defaultServices = [
        'M365 Teams',
        'M365 Email',
        'Software Distribution (SCCM)',
        'M365 Apps',
        'Ask IT',
        'EMEIA Messaging',
        'Mobile Phones UK',
        'ZinZai Connect',
        'ForcePoint',
        'Network Service (CE/WEMEIA)',
        'M365 Sharepoint'
    ];

    const [randomServices, setRandomServices] = useState(services.reduce((acc, service) => ({ ...acc, [service]: defaultServices.includes(service) }), {}));

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
        if (e.target.files.length > 0) {
            const file = e.target.files[0];
            document.getElementById('file-name').innerHTML = ` <note class="remove_bold">Selected file:</note> ${file.name}`;
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
        }
    };

    const handleConfigSubmit = async () => {
        const sfMembersArray = sfMembers.split('\n').map(name => name.trim()).filter(name => name.length > 0);

        try {
            const response = await fetch(processUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    ...config,
                    sfMembers: sfMembersArray,
                    agentNames: agentAccounts.split('\n'),
                    randomServices: Object.keys(randomServices).filter(service => randomServices[service])
                }),
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
        "Kempa, Martin\nSocha, Michał\nKrasowicz, Barbara\nSzczypior, Dawid\nSiemieniuk, Roman\nKalbarczyk, Jan\nLubonski, Piotr\nKucinska, Diana\nStepien, Ewa\nKoplin, Krzysztof"
    );
    const [agentAccounts, setAgentAccounts] = useState('');

    return (
        <div>

            <div class="block-1">
                <div class="section-number-1">1</div>
                <div class="section-content-1">
                    <tit-2>Upload Excel file report</tit-2>
                    <br />
                    <label className="custom-file-upload">
                        <input type="file" ref={fileInput} onChange={handleFileInput} style={{ display: 'none' }} />
                        <span>Upload File</span>
                        <span id="file-name"></span>
                    </label>
                    <br />
                </div>
            </div>

            <hr></hr>

            <div class="block-2">
                <div class="section-number-2">2</div>
                <div class="section-content-2">
                    <tit-2>Incidents per Agent</tit-2>
                    <br />
                    <select name="incidentsPerAgent" value={config.incidentsPerAgent} onChange={(e) => setConfig({ ...config, incidentsPerAgent: parseInt(e.target.value) })}>
                        {[5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15].map(num => <option key={num} value={num}>{num}</option>)}
                    </select>
                </div>
            </div>

            <hr></hr>

            <div class="block-3">
                <div class="section-number-3">3</div>
                <div class="section-content-3">
                    <div className="config-random-grid">
                        <div className="services-config-grid">
                            <strong></strong>
                            <tit-2>Services</tit-2>
                            {/* <strong>Contact Type:</strong> <strong>FTF:</strong> */}
                            {config.incidentConfigs.map((incidentConfig, index) => (
                                <React.Fragment key={index}>
                                    <span class="list-bullets">{index + 1}</span>
                                    <span class="services-config">
                                    <select value={incidentConfig.service} onChange={(e) => handleIncidentChange(index, 'service', e.target.value)}>
                                        {services.map(service => <option key={service} value={service}>{service}</option>)}
                                    </select>
                                    {/* <select value={incidentConfig.contactType} onChange={(e) => handleIncidentChange(index, 'contactType', e.target.value)}> {contactTypes.map(type => <option key={type} value={type}>{type}</option>)} </select> <select value={incidentConfig.ftf} onChange={(e) => handleIncidentChange(index, 'ftf', e.target.value)}> <option value="TRUE">TRUE</option> <option value="FALSE">FALSE</option> <option value="RANDOM">RANDOM</option> </select> */}
                                    </span>
                                </React.Fragment>
                            ))}
                        </div>
                       {/*  <br /> */}

                        <vertical-line></vertical-line>

                        <div className='random-services'>
                            <tit-2>Define 'RANDOM'</tit-2>
                            {services.filter(service => service !== 'RANDOM').map(service => (
                                <label key={service}>
                                    <input
                                        type="checkbox"
                                        checked={randomServices[service]}
                                        onChange={() => setRandomServices({ ...randomServices, [service]: !randomServices[service] })}
                                    />
                                    {service}
                                </label>
                            ))}
                        </div>
                    </div>
                </div>
            </div>

            <hr></hr>

            <div class="block-4">
                <div class="section-number-4">4</div>
                <div class="section-content-4">
                    <div className='all-accounts'>
                        <div className='sf-accounts'>
                            <tit-2>SF team members</tit-2>
                            <br />
                            Change accounts manually <br /><br />
                            <textarea
                                rows="10"
                                cols="30"
                                value={sfMembers}
                                onChange={e => setSfMembers(e.target.value)}
                                placeholder="Paste the list of SF team members here, one per line."
                            />
                        </div>

                        {/* <vertical-line></vertical-line> */}

                        <div className='agent-accounts'>
                            <tit-2>Agent Accounts</tit-2>
                            <br />
                            Delete unwanted accounts <br /><br />
                            <textarea
                                rows="10"
                                value={agentAccounts}
                                onChange={e => setAgentAccounts(e.target.value)}
                                placeholder="Agent accounts will be automatically populated here after file upload."
                            />
                        </div>
                    </div>
                </div>
            </div>

            <hr></hr>

            <div class="block-5">
                <div class="section-number-5">5</div>
                <div class="section-content-5">
                    <tit-2>Process and Download</tit-2>
                    <br />
                    <input type="button" value="Process" className="process-button" onClick={handleConfigSubmit} />
                    <br /><br />
                </div >
            </div>
        </div >

    );
};

export default FileUpload;