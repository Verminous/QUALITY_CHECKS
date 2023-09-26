const express = require('express');
const multer = require('multer');
const app = express();
const port = 3001;

const upload = multer({ dest: 'uploads/' });

app.get('/', (req, res) => {
  res.send('Server running!');
});

app.post('/upload', upload.single('file'), (req, res) => {
  console.log('File received:', req.file);
  
  // TODO: Process the file here
  
  res.json({ message: 'File received' });
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/`);
});


