const express = require('express');
const router = express.Router();
const exceljs = require('exceljs');
const xlsxPopulate = require('xlsx-populate');

function moveCellsDown(sheet, startRow, numRows) {
  // Переносим строки вниз начиная с конца
  for (let i = sheet.rowCount; i >= startRow; i--) {
    // Получаем исходную строку
    const sourceRow = sheet.getRow(i);

    // Создаем новую строку для переноса
    const targetRowIndex = i + numRows;
    let targetRow = sheet.getRow(targetRowIndex);
    if (!targetRow) {
      // Если строки не существует, создаем новую строку
      sheet.spliceRows(targetRowIndex, 0, [{}]);
      targetRow = sheet.getRow(targetRowIndex);
    }

    // Переносим каждую ячейку в строке
    sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = targetRow.getCell(colNumber);

      // Проверяем, является ли ячейка объединенной
      if (cell.isMerged) {
        // Находим главную (первую) ячейку в объединенной группе
        const masterCell = sheet.getCell(cell.master.address);
        if (masterCell === cell) {
          // Копируем значение главной ячейки в целевую ячейку
          targetCell.value = masterCell.value;
          targetCell.style = masterCell.style;
          targetCell.numFmt = masterCell.numFmt;
          targetCell.border = masterCell.border;
        }
      } else {
        // Копируем значение ячейки в целевую ячейку
        targetCell.value = cell.value;
        targetCell.style = cell.style;
        targetCell.numFmt = cell.numFmt;
        targetCell.border = cell.border;
      }
    });
  }
}

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

      // Пример использования функции
      // moveCellsDown(sheet, rowIndex + 1, 13);
      const newRow1Values = [
        '5.1.1',
        tableData.winner,
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData1,
      ];
      const newRow2Values = ['', '', '', '2-х мест', tableData.commandData2];
      const newRow3Values = ['', '', '', '3-х мест', tableData.commandData3];
      const newRow4Values = [
        '',
        '',
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData11,
      ];
      const newRow5Values = ['', '', '', '2-х мест', tableData.commandData21];
      const newRow6Values = ['', '', '', '3-х мест', tableData.commandData31];
      const newRow7Values = [
        '',
        '',
        'Призовые места по итогам личного первенства в номинациях',
        '1-х мест',
        tableData.personalData1,
      ];
      const newRow8Values = ['', '', '', '2-х мест', tableData.personalData2];
      const newRow9Values = ['', '', '', '3-х мест', tableData.personalData3];
      const newRow10Values = ['', '', 'Гран При', '', tableData.grandPrizeData];
      const newRow11Values = [
        '',
        '',
        'Приз за отдельные достижения',
        '',
        tableData.individualAchievementData,
      ];
      const newRow12Values = [
        '',
        '',
        'Побед в специальных номинациях',
        '',
        tableData.specialAwardsData,
      ];
      const newRow13Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '',
        tableData.lackOfCompetitiveComponentData,
      ];

      // Форматирование ячейки A1

      sheet.spliceRows(rowIndex + 1, 0, newRow1Values);
      sheet.spliceRows(rowIndex + 2, 0, newRow2Values);
      sheet.spliceRows(rowIndex + 3, 0, newRow3Values);
      sheet.spliceRows(rowIndex + 4, 0, newRow4Values);
      sheet.spliceRows(rowIndex + 5, 0, newRow5Values);
      sheet.spliceRows(rowIndex + 6, 0, newRow6Values);
      sheet.spliceRows(rowIndex + 7, 0, newRow7Values);
      sheet.spliceRows(rowIndex + 8, 0, newRow8Values);
      sheet.spliceRows(rowIndex + 9, 0, newRow9Values);
      sheet.spliceRows(rowIndex + 10, 0, newRow10Values);
      sheet.spliceRows(rowIndex + 11, 0, newRow11Values);
      sheet.spliceRows(rowIndex + 12, 0, newRow12Values);
      sheet.spliceRows(rowIndex + 13, 0, newRow13Values);
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 13;

      for (let i = startRowIndex; i <= endRowIndex; i++) {
        for (let j = startRowIndex; j <= endRowIndex; j++) {
          if (i !== j) {
            sheet.unMergeCells(`A${i}:A${j}`);
            sheet.unMergeCells(`B${i}:B${j}`);
            sheet.unMergeCells(`C${i}:C${j}`);
            sheet.unMergeCells(`D${i}:D${j}`);
            sheet.unMergeCells(`E${i}:E${j}`);
            sheet.unMergeCells(`F${i}:F${j}`);
            sheet.unMergeCells(`G${i}:G${j}`);
          }
        }
      }

      sheet.mergeCells(`A${rowIndex + 1}:A${rowIndex + 13}`);
      sheet.mergeCells(`B${rowIndex + 1}:B${rowIndex + 13}`);
      sheet.mergeCells(`C${rowIndex + 1}:C${rowIndex + 3}`);
      sheet.mergeCells(`C${rowIndex + 4}:C${rowIndex + 6}`);
      sheet.mergeCells(`C${rowIndex + 7}:C${rowIndex + 9}`);
      sheet.mergeCells(`C${rowIndex + 10}:D${rowIndex + 10}`);
      sheet.mergeCells(`C${rowIndex + 11}:D${rowIndex + 11}`);
      sheet.mergeCells(`C${rowIndex + 12}:D${rowIndex + 12}`);
      sheet.mergeCells(`C${rowIndex + 13}:D${rowIndex + 13}`);

      for (let i = rowIndex + 1; i <= endRowIndex; i++) {
        const cellA = sheet.getCell(`A${i}`);
        const cellB = sheet.getCell(`B${i}`);
        const cellC = sheet.getCell(`C${i}`);
        const cellD = sheet.getCell(`D${i}`);
        const cellE = sheet.getCell(`E${i}`);
        const cellF = sheet.getCell(`F${i}`);
        const cellG = sheet.getCell(`G${i}`);

        // Применяем стили к ячейке A
        cellA.style.font = {
          name: 'Times New Roman',
          size: 12,
        };
        cellA.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };

        // Применяем стили к ячейке B
        cellB.style.font = {
          italic: true,
          name: 'Times New Roman',
          size: 12,
        };
        cellB.alignment = {
          vertical: 'top',
          horizontal: 'left',
          wrapText: true,
        };
        cellB.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };

        // Применяем стили к ячейке C
        cellC.style.font = {
          name: 'Times New Roman',
          size: 12,
        };
        cellC.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };

        // Применяем стили к ячейкам D, E, F, G
        [cellD, cellE, cellF, cellG].forEach((cell) => {
          cell.style.font = {
            name: 'Times New Roman',
            size: 12,
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      }

      const currentDate = new Date();
      const dateString = currentDate.toISOString().slice(0, 10); // Преобразовать текущую дату в строку формата "гггг-мм-дд"
      const timeString = currentDate.toTimeString().slice(0, 8).replace(/:/g, '-'); // Преобразовать текущее время в строку и удалить двоеточия, заменив их на дефисы
      const fileName = `table_${dateString}_${timeString}.xlsx`; // Имя файла с добавленной датой и временем

      await workbook.xlsx.writeFile(fileName);

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
