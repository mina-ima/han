import express from 'express';

const app = express();
const port = 3002;

app.get('/', (req, res) => {
  res.send('Hello from Backend!');
});

app.listen(port, () => {
  console.log(`Backend server listening at http://localhost:${port}`);
});