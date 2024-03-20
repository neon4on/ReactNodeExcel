const express = require('express');
const app = express();
const cors = require('cors');
const bodyParser = require('body-parser');
const excelRoutes = require('./excelRoutes');

const port = 4000;

app.use(cors());
app.use(bodyParser.json());

app.use('/api', excelRoutes);

app.listen(port, () => {
  console.log(`Server is listening at http://localhost:${port}`);
});

module.exports = app;
