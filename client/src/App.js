import React from 'react';
import FileUpload from './FileUpload';
import './App.css';

function App() {
  const handleFileSelect = (file) => {
    console.log("Selected file:", file);
    // You can now send this file to your server or handle it according to your needs
  };

  const handleConfigSubmit = (config) => {
    console.log("Configurations:", config);
    // Send this configuration to your server along with the file
  };

  return (
    <div className="my-uploader">
      <FileUpload onFileSelect={handleFileSelect} onConfigSubmit={handleConfigSubmit} />
    </div>
  );
}

export default App;
