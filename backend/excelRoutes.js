const express = require('express');
const router = express.Router();
const exceljs = require('exceljs');
const xlsxPopulate = require('xlsx-populate');

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
function columnIndexToName(index) {
  let name = '';
  while (index > 0) {
    const remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - remainder) / 26);
  }
  return name;
}

function columnNameToIndex(name) {
  let index = 0;
  for (let i = 0; i < name.length; i++) {
    index = index * 26 + name.charCodeAt(i) - 64;
  }
  return index;
}

router.post('/createExcel54', async (req, res) => {
  const tableData = req.body;

  async function insertDataAfterRow(filePath, sheetName, columnAValue, columnBValue) {
    try {
      const workbook = await xlsxPopulate.fromFileAsync(filePath);
      const sheet = workbook.sheet(sheetName);

      // Находим индекс строки по значению в столбце A
      let rowIndex = null;
      let columnIndexB = null;
      sheet
        .usedRange()
        .value()
        .forEach((row, index) => {
          if (row[0] === columnAValue) {
            rowIndex = index;
            console.log(`Значение ${columnAValue} найдено в столбце A.`);
          }
          if (row.includes(columnBValue)) {
            columnIndexB = row.indexOf(columnBValue);
            console.log(`Значение ${columnBValue} найдено в строке.`);
          }
        });

      if (rowIndex !== null && columnIndexB !== null) {
        // Вставляем новые данные в строку следующую за найденной
        const newRow = sheet.row(rowIndex + 2);

        // Значения для столбца B
        newRow.cell(columnIndexB + 1).value(tableData.winner);

        // Значения для столбца C
        newRow.cell(columnIndexB + 1).value('Побед');
        newRow.cell(columnIndexB + 2).value('Призовых мест');
        newRow.cell(columnIndexB + 3).value('Отсутствие соревновательной составляющей');

        // Значения для столбца E
        newRow.cell(columnIndexB + 5).value(tableData.commandData1);
        newRow.cell(columnIndexB + 6).value(tableData.commandData2);
        newRow.cell(columnIndexB + 7).value(tableData.commandData3);

        // Сохраняем изменения в файл
        await workbook.toFileAsync(filePath);

        console.log('Данные успешно вставлены после строки с значением', columnAValue);
      } else {
        console.log(`Значение ${columnAValue} или ${columnBValue} не найдено.`);
      }
    } catch (error) {
      console.error('Ошибка при вставке данных:', error);
    }
  }

  // Пример использования функции
  const filePath = 'example.xlsx';
  const sheetName = 'ЦОВУ';
  const columnAValue = '5.4';
  const columnBValue = tableData.select;
  console.log(`Значение ${tableData.select} это B`);
  await insertDataAfterRow(filePath, sheetName, columnAValue, columnBValue);
});

router.post('/createExcel64', (req, res) => {
  const tableData = req.body;

  const workbook = new exceljs.Workbook();
  const sheet = workbook.addWorksheet('Sheet 1');

  sheet.mergeCells('B1', 'D1');

  sheet.getCell('A1').value = tableData.select;

  sheet.getCell('B1').value = tableData.winner;

  sheet.getCell('E1').value = tableData.commandData1;

  workbook.xlsx
    .writeFile('table64.xlsx')
    .then(() => {
      console.log('File saved successfully');
      res.send('File saved successfully');
    })
    .catch((error) => {
      console.error('Error saving file:', error);
      res.status(500).send('Error saving file');
    });
});

router.post('/createExcel723', (req, res) => {
  const tableData = req.body;

  const workbook = new exceljs.Workbook();
  const sheet = workbook.addWorksheet('Sheet 1');

  sheet.mergeCells('B1', 'B7');
  sheet.mergeCells('C1', 'C3');
  sheet.mergeCells('C4', 'C6');
  sheet.mergeCells('C7', 'D7');

  sheet.getCell('B1').value = tableData.winner;

  sheet.getCell('C1').value = 'Призовые места по итогам командного первенства в номинациях';
  sheet.getCell('C4').value = 'Призовые места по итогам личного первенства в номинациях';
  sheet.getCell('C7').value = 'Отсутствие соревновательной составляющей';

  sheet.getCell('D1').value = '1-х мест';
  sheet.getCell('D2').value = '2-х мест';
  sheet.getCell('D3').value = '3-х мест';
  sheet.getCell('D4').value = '1-х мест';
  sheet.getCell('D5').value = '2-х мест';
  sheet.getCell('D6').value = '3-х мест';

  sheet.getCell('E1').value = tableData.commandData1;
  sheet.getCell('E2').value = tableData.commandData2;
  sheet.getCell('E3').value = tableData.commandData3;
  sheet.getCell('E4').value = tableData.personalData1;
  sheet.getCell('E5').value = tableData.personalData2;
  sheet.getCell('E6').value = tableData.personalData3;
  sheet.getCell('E7').value = tableData.lackOfCompetitiveComponentData;

  workbook.xlsx
    .writeFile('table723.xlsx')
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
