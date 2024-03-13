const express = require('express');
const router = express.Router();
const exceljs = require('exceljs');
const xlsxPopulate = require('xlsx-populate');

router.post('/createExcel51', async function (req, res) {
  const tableData = req.body;

  try {
    const workbook = new exceljs.Workbook();
    const filePath = 'example.xlsx'; // Путь к файлу

    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet('ЦОВУ'); // Получение доступа к странице "ЦОВУ"

    if (sheet) {
      console.log('Страница "ЦОВУ" найдена.');

      // Ваша логика работы с листом
    } else {
      console.log('Страница "ЦОВУ" не найдена.');
    }

    // Находим индекс строки по значению в столбце A
    let rowIndex = null;

    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      // console.log(`Значение в строке ${index}: ${rowData}`);

      if (rowData == 5.1 || rowData == '5,1' || rowData == '5.1') {
        rowIndex = index;
        console.log(`Значение 5.1 найдено в строке ${index}.`);
      }
    });

    if (rowIndex !== null) {
      console.log(`Строка с 5.1 найдена: ${rowIndex}`);

      function moveCellsDown(sheet, startRow, numRows) {
        // Переносим строки вниз начиная с конца
        for (let i = sheet.rowCount; i >= startRow; i--) {
          // Получаем строку
          const sourceRow = sheet.getRow(i);

          // Создаем новую строку для переноса
          const targetRowNumber = i + numRows;
          let targetRow = sheet.getRow(targetRowNumber);
          if (!targetRow) {
            // Если строки не существует, создаем новую строку
            sheet.spliceRows(targetRowNumber, 0, [{}]);
            targetRow = sheet.getRow(targetRowNumber);
          }
          targetRow.hidden = sourceRow.hidden;

          // Переносим каждую ячейку в строке
          sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const targetCell = targetRow.getCell(colNumber);

            // Проверяем наличие ячейки
            if (cell && targetCell) {
              // Копируем только главные ячейки объединенной группы
              if (cell.isMerged && cell.master === cell) {
                // Копируем данные из главной ячейки
                targetCell.value = cell.value;
                targetCell.style = cell.style;
                targetCell.numFmt = cell.numFmt;
                targetCell.border = cell.border;
              }
            }
          });

          // Очищаем исходную строку
          // sourceRow.eachCell({ includeEmpty: true }, (cell) => {
          //     cell.value = null;
          //     cell.style = {};
          //     cell.numFmt = null;
          //     cell.border = {};
          // });
        }
      }

      // Пример использования функции
      moveCellsDown(sheet, rowIndex + 1, 13);

      // Пример использования функции
      moveCellsDown(sheet, rowIndex + 1, 13);

      // Форматирование ячейки A1
      const cellA1 = sheet.getCell(`A${rowIndex + 1}`);
      cellA1.value = '5.1.1';
      // cellA1.numFmt = '0.00';
      cellA1.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      // Значения для столбца B
      sheet.getCell(`B${rowIndex + 1}`).value = tableData.winner;
      sheet.getCell(`B${rowIndex + 1}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      sheet.getCell(`B${rowIndex + 1}`).alignment = { vertical: 'middle', horizontal: 'center' };
      // Значения для столбца C
      sheet.getCell(`C${rowIndex + 1}`).value =
        'Призовые места по итогам командного первенства в номинациях';
      sheet.getCell(`C${rowIndex + 4}`).value =
        'Призовые места по итогам командного первенства в номинациях';
      sheet.getCell(`C${rowIndex + 7}`).value =
        'Призовые места по итогам личного первенства в номинациях';
      sheet.getCell(`C${rowIndex + 10}`).value = 'Гран При';
      sheet.getCell(`C${rowIndex + 11}`).value = 'Приз за отдельные достижения';
      sheet.getCell(`C${rowIndex + 12}`).value = 'Побед в специальных номинациях';
      sheet.getCell(`C${rowIndex + 13}`).value = 'Отсутствие соревновательной составляющей';

      sheet.getCell(`D${rowIndex + 1}`).value = '1-х мест';
      sheet.getCell(`D${rowIndex + 2}`).value = '2-х мест';
      sheet.getCell(`D${rowIndex + 3}`).value = '3-х мест';
      sheet.getCell(`D${rowIndex + 4}`).value = '1-х мест';
      sheet.getCell(`D${rowIndex + 5}`).value = '2-х мест';
      sheet.getCell(`D${rowIndex + 6}`).value = '3-х мест';
      sheet.getCell(`D${rowIndex + 7}`).value = '1-х мест';
      sheet.getCell(`D${rowIndex + 8}`).value = '2-х мест';
      sheet.getCell(`D${rowIndex + 9}`).value = '3-х мест';

      // Значения для столбца E
      sheet.getCell(`E${rowIndex + 1}`).value = tableData.commandData1;
      sheet.getCell(`E${rowIndex + 2}`).value = tableData.commandData2;
      sheet.getCell(`E${rowIndex + 3}`).value = tableData.commandData3;
      sheet.getCell(`E${rowIndex + 4}`).value = tableData.commandData11;
      sheet.getCell(`E${rowIndex + 5}`).value = tableData.commandData21;
      sheet.getCell(`E${rowIndex + 6}`).value = tableData.commandData31;
      sheet.getCell(`E${rowIndex + 7}`).value = tableData.personalData1;
      sheet.getCell(`E${rowIndex + 8}`).value = tableData.personalData2;
      sheet.getCell(`E${rowIndex + 9}`).value = tableData.personalData3;
      sheet.getCell(`E${rowIndex + 10}`).value = tableData.grandPrizeData;
      sheet.getCell(`E${rowIndex + 11}`).value = tableData.individualAchievementData;
      sheet.getCell(`E${rowIndex + 12}`).value = tableData.specialAwardsData;
      sheet.getCell(`E${rowIndex + 13}`).value = tableData.lackOfCompetitiveComponentData;

      await workbook.xlsx.writeFile(filePath);

      console.log('Данные успешно вставлены после строки с значением 5.1.');
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 5.1 не найдена.');
    }
  } catch (error) {
    console.error('Ошибка при создании Excel файла:', error);
  }
});

router.post('/createExcel54', async (req, res) => {
  const tableData = req.body;
  console.dir(req.body);

  const arr = [
    'Муниципальных соревнований',
    'Региональных соревнований',
    'Всероссийских соревнований',
    'Международных соревнований',
  ];

  let valueShort = null;
  let flagShort = false;
  async function insertDataAfterRow(filePath, sheetName, columnAValue, columnBValue) {
    try {
      const workbook = await xlsxPopulate.fromFileAsync(filePath);
      const sheet = workbook.sheet(sheetName);

      // Находим индекс строки по значению в столбце A
      let rowIndex = null;
      let columnIndexB = null;

      for (let i = 0; i <= arr.length; i++) {
        if (columnBValue == arr[i]) {
          if (i < arr.length - 1) {
            valueShort = arr[i + 1];
            flagShort = true;
          } else {
            valueShort = arr[i];
          }
        }
      }

      sheet
        .usedRange()
        .value()
        .forEach((row, index) => {
          if (row.includes(6.4) || row.includes('6,4') || row.includes('6.4')) {
            rowIndex = index + 1;
            console.log(`Значение ${columnAValue} найдено в столбце A.`);
          }
          if (row.includes(columnBValue)) {
            columnIndexB = index + 1;
            console.log(`Значение ${columnBValue} найдено в строке.`);
          }

          if (row.includes(valueShort)) {
            valueShort = index + 1;
            console.log(`Значение ${valueShort} найдено в строке.`);
          }
        });

      if (rowIndex !== null && columnIndexB !== null && valueShort !== null) {
        // Вставляем новые данные в строку следующую за найденной

        console.log(rowIndex + ' A');
        console.log(columnIndexB + ' B');
        console.log(valueShort + ' Short');
        const value = sheet.usedRange().value().length;
        console.log(value + ' END');

        // for (let i = value; i >= columnIndexB; i--) {
        //   const sourceRow = sheet.row(i);
        //   const targetRow = sheet.row(i + 3);
        //   // Копирование стилей ячеек
        //   // Копирование стилей ячеек
        //   for (let j = 1; j <= value; j++) {
        //     const sourceCell = sourceRow.cell(j);
        //     const targetCell = targetRow.cell(j);
        //     const sourceStyle = sourceCell.style();
        //     targetCell.style({
        //       fontName: sourceStyle.fontName(),
        //       bold: sourceStyle.bold(),
        //       italic: sourceStyle.italic(),
        //       // Добавьте другие свойства стиля по необходимости
        //     });
        //   }
        // }
        sheet.getColumn('A').eachCell({ includeEmpty: true }, (cell) => {
          cell.style({ numberFormat: '0.00' });
        });
        // Значения для столбца B
        sheet.cell(`B${columnIndexB + 1}`).value(tableData.winner);

        // Значения для столбца C
        sheet.cell(`C${columnIndexB + 1}`).value('Побед');
        sheet.cell(`C${columnIndexB + 2}`).value('Призовых мест');
        sheet.cell(`C${columnIndexB + 3}`).value('Отсутствие соревновательной составляющей');

        // Значения для столбца E
        sheet.cell(`E${columnIndexB + 1}`).value(tableData.commandData1);
        sheet.cell(`E${columnIndexB + 2}`).value(tableData.commandData2);
        sheet.cell(`E${columnIndexB + 3}`).value(tableData.commandData3);

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
  const columnAValue = 5.4;
  const columnBValue = tableData.select;
  await insertDataAfterRow(filePath, sheetName, columnAValue, columnBValue);
});

router.post('/createExcel64', async (req, res) => {
  // const tableData = req.body;

  // const workbook = new exceljs.Workbook();
  // const sheet = workbook.addWorksheet('Sheet 1');

  // // sheet.mergeCells('B1', 'D1');

  // sheet.getCell('A1').value = tableData.select;

  // sheet.getCell('B1').value = tableData.winner;

  // sheet.getCell('E1').value = tableData.commandData1;
  const tableData = req.body;
  console.dir(req.body);

  const arr = ['Регионального уровня', 'Всероссийского уровня'];

  let valueShort = null;
  let flagShort = false;
  async function insertDataAfterRow(filePath, sheetName, columnAValue, columnBValue) {
    try {
      const workbook = await xlsxPopulate.fromFileAsync(filePath);
      const sheet = workbook.sheet(sheetName);

      // Находим индекс строки по значению в столбце A
      let rowIndex = null;
      let columnIndexB = null;

      for (let i = 0; i < arr.length; i++) {
        if (columnBValue == arr[i]) {
          if (i < arr.length - 1) {
            valueShort = arr[i + 1];
            flagShort = true;
          } else {
            valueShort = arr[i];
          }
        }
      }

      sheet
        .usedRange()
        .value()
        .forEach((row, index) => {
          if (row.includes(6.4) || row.includes('6,4') || row.includes('6.4')) {
            rowIndex = index + 1;
            console.log(`Значение ${columnAValue} найдено в столбце A.`);
          }
          if (row.includes(columnBValue)) {
            columnIndexB = index + 1;
            console.log(`Значение ${columnBValue} найдено в строке.`);
          }

          if (row.includes(valueShort)) {
            valueShort = index + 1;
            console.log(`Значение ${valueShort} найдено в строке.`);
          }
        });

      if (rowIndex !== null && columnIndexB !== null && valueShort !== null) {
        // Вставляем новые данные в строку следующую за найденной
        console.log(rowIndex + ' A');
        console.log(columnIndexB + ' B');
        console.log(valueShort + ' Short');
        const value = sheet.usedRange().value().length;
        console.log(value + ' END');

        // Применение стиля к столбцу A
        // Применение стиля к столбцу A
        const usedRange = sheet.usedRange();
        for (let rowIndex = 0; rowIndex < usedRange.rowCount; rowIndex++) {
          const cell = sheet.cell(rowIndex + 1, 1);
          cell.style({ numberFormat: '0.0' });
        }

        for (let rowIndex = 0; rowIndex < usedRange.rowCount; rowIndex++) {
          const cellValue = sheet.cell(rowIndex + 1, 1).value();
          if (typeof cellValue === 'number') {
            // Если значение ячейки является числом, заменяем запятую на точку
            const updatedValue = cellValue.toString().replace(',', '.');
            sheet.cell(rowIndex + 1, 1).value(updatedValue);
          }
        }

        // Применяем стили и форматирование для чисел

        // Значения для столбца B
        sheet.cell(`B${columnIndexB + 1}`).value(tableData.winner);

        // Значения для столбца C
        sheet.cell(`C${columnIndexB + 1}`).value('Побед');
        sheet.cell(`C${columnIndexB + 2}`).value('Призовых мест');
        sheet.cell(`C${columnIndexB + 3}`).value('Отсутствие соревновательной составляющей');

        // Значения для столбца E
        sheet.cell(`E${columnIndexB + 1}`).value(tableData.commandData1);
        sheet.cell(`E${columnIndexB + 2}`).value(tableData.commandData2);
        sheet.cell(`E${columnIndexB + 3}`).value(tableData.commandData3);

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

  const filePath = 'example.xlsx';
  const sheetName = 'ЦОВУ';
  const columnAValue = 6.4;
  const columnBValue = tableData.select;
  await insertDataAfterRow(filePath, sheetName, columnAValue, columnBValue);
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
