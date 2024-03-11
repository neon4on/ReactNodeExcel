const express = require('express');
const router = express.Router();
const exceljs = require('exceljs');

router.post('/createExcel51', (req, res) => {
  const tableData = req.body;

  const workbook = new exceljs.Workbook();
  const sheet = workbook.addWorksheet('Sheet 1');

  sheet.mergeCells('A1', 'A13');
  sheet.mergeCells('B1', 'B13');
  sheet.mergeCells('C1', 'C3');
  sheet.mergeCells('C4', 'C6');
  sheet.mergeCells('C7', 'C9');
  sheet.mergeCells('C10', 'D10');
  sheet.mergeCells('C11', 'D11');
  sheet.mergeCells('C12', 'D12');
  sheet.mergeCells('C13', 'D13');

  sheet.getCell('A1').value = '5.1.1';

  sheet.getCell('B1').value = tableData.winner;

  sheet.getCell('C1').value = 'Призовые места по итогам командного первенства в номинациях';
  sheet.getCell('C4').value = 'Призовые места по итогам командного первенства в номинациях';
  sheet.getCell('C7').value = 'Призовые места по итогам личного первенства в номинациях';
  sheet.getCell('C10').value = 'Гран При';
  sheet.getCell('C11').value = 'Приз за отдельные достижения';
  sheet.getCell('C12').value = 'Побед в специальных номинациях';
  sheet.getCell('C13').value = 'Отсутствие соревновательной составляющей';

  sheet.getCell('D1').value = '1-х мест';
  sheet.getCell('D2').value = '2-х мест';
  sheet.getCell('D3').value = '3-х мест';
  sheet.getCell('D4').value = '1-х мест';
  sheet.getCell('D5').value = '2-х мест';
  sheet.getCell('D6').value = '3-х мест';
  sheet.getCell('D7').value = '1-х мест';
  sheet.getCell('D8').value = '2-х мест';
  sheet.getCell('D9').value = '3-х мест';
  sheet.getCell('D9').value = '3-х мест';

  sheet.getCell('E1').value = tableData.commandData1;
  sheet.getCell('E2').value = tableData.commandData2;
  sheet.getCell('E3').value = tableData.commandData3;
  sheet.getCell('E4').value = tableData.commandData11;
  sheet.getCell('E5').value = tableData.commandData21;
  sheet.getCell('E6').value = tableData.commandData31;
  sheet.getCell('E7').value = tableData.personalData1;
  sheet.getCell('E8').value = tableData.personalData2;
  sheet.getCell('E9').value = tableData.personalData3;
  sheet.getCell('E10').value = tableData.grandPrizeData;
  sheet.getCell('E11').value = tableData.individualAchievementData;
  sheet.getCell('E12').value = tableData.specialAwardsData;
  sheet.getCell('E13').value = tableData.lackOfCompetitiveComponentData;

  workbook.xlsx
    .writeFile('table51.xlsx')
    .then(() => {
      console.log('File saved successfully');
      res.send('File saved successfully');
    })
    .catch((error) => {
      console.error('Error saving file:', error);
      res.status(500).send('Error saving file');
    });
});

router.post('/createExcel54', (req, res) => {
  const tableData = req.body;

  const workbook = new exceljs.Workbook();
  const sheet = workbook.addWorksheet('Sheet 1');

  sheet.mergeCells('B1', 'B3');
  sheet.mergeCells('C1', 'D1');
  sheet.mergeCells('C2', 'D2');
  sheet.mergeCells('C3', 'D3');

  sheet.getCell('A1').value = tableData.select;

  sheet.getCell('B1').value = tableData.winner;

  sheet.getCell('C1').value = 'Побед';
  sheet.getCell('C2').value = 'Призовых мест';
  sheet.getCell('C3').value = 'Отсутствие соревновательной составляющей';

  sheet.getCell('E1').value = tableData.commandData1;
  sheet.getCell('E2').value = tableData.commandData2;
  sheet.getCell('E3').value = tableData.commandData3;

  workbook.xlsx
    .writeFile('table54.xlsx')
    .then(() => {
      console.log('File saved successfully');
      res.send('File saved successfully');
    })
    .catch((error) => {
      console.error('Error saving file:', error);
      res.status(500).send('Error saving file');
    });
});

module.exports = router;
