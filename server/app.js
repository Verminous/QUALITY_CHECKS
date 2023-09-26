const express = require('express');
const app = express();
const path = require('path');
const port = 3001; // Note: using a different port from React


app.get('/api/hello', (req, res) => {
  res.send({ message: 'Hello from server!' });
});

// app.get('/', (req, res) => {
//   res.send('Hello, world!');
// });

// Serve static files from the React app
app.use(express.static(path.join(__dirname, '../client/build')));

// The "catchall" handler: for any request that doesn't
// match one above, send back React's index.html file.
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname + '../client/build/index.html'));
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/`);
});
