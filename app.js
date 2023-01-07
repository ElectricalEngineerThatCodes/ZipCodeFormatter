const { getWorkbook } = require('./ZipFormatter.js');
const express = require('express');
const app = express();
const port = 3000;

app.get('/', (req, res) => {
  res.send('Hello World!');
});

document.querySelector('.check').addEventListener('click', function () {
  app.get('/donors', async (req, res) => {
    const fileName = 'FormatForMailing2.xlsx';

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    res.setHeader('Content-Disposition', 'attachment; filename=' + fileName);

    const workbook = await getWorkbook();

    await workbook.xlsx.write(res);
    res.status(200).end();
    //   res.end();
  });
});

app.get('/users/atown', (req, res) => {
  res.send('Andrew Townley');
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
