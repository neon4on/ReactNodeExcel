const express = require('express');
const router = express.Router();
const exceljs = require('exceljs');
const xlsxPopulate = require('xlsx-populate');

router.post('/createExcel51', async function (req, res) {
  const tableData = req.body;

  try {
    const workbook = new exceljs.Workbook();
    const filePath = 'example.xlsx';

    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet('Лист1');

    if (sheet) {
      console.log('Страница "Лист1" найдена.');
    } else {
      console.log('Страница "Лист1" не найдена.');
    }

    let rowIndex = null;
    let rowEndIndex = null;

    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData == 5.1 || rowData == '5,1' || rowData == '5.1') {
        rowIndex = index;
        console.log(`Значение 5.1 найдено в строке ${index}.`);
      }
      if (rowData == 5.2 || rowData == '5,2' || rowData == '5.2') {
        rowEndIndex = index;
        console.log(`Значение 5.2 найдено в строке ${index}.`);
      }
    });

    if (rowIndex !== null) {
      console.log(`Строка с 5.1 найдена: ${rowIndex}`);

      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 13;

      let j = 0;
      for (let i = startRowIndex + 1; i < rowEndIndex; i++) {
        if (j % 13 == 0) {
          const rowData = sheet.getCell(`A${i}`).value;
          const parts = rowData.split('.');
          const number = parseInt(parts[2]);
          parts[2] = (number + 1).toString();
          sheet.getCell(`A${i}`).value = parts.join('.');
        }
        j++;
      }

      const newRow1Values = [
        '5.1.1',
        tableData.winner,
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData1,
        tableData.commandData1 * 30,
      ];
      const newRow2Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData2,
        tableData.commandData2 * 25,
      ];
      const newRow3Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData3,
        tableData.commandData3 * 20,
      ];
      const newRow4Values = [
        '',
        '',
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData11,
        tableData.commandData11 * 10,
      ];
      const newRow5Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData21,
        tableData.commandData21 * 8,
      ];
      const newRow6Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData31,
        tableData.commandData31 * 6,
      ];
      const newRow7Values = [
        '',
        '',
        'Призовые места по итогам личного первенства в номинациях',
        '1-х мест',
        tableData.personalData1,
        tableData.personalData1 * 5,
      ];
      const newRow8Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.personalData2,
        tableData.personalData2 * 4,
      ];
      const newRow9Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.personalData3,
        tableData.personalData3 * 3,
      ];
      const newRow10Values = [
        '',
        '',
        'Гран При',
        '',
        tableData.grandPrizeData,
        tableData.grandPrizeData * 15,
      ];
      const newRow11Values = [
        '',
        '',
        'Приз за отдельные достижения',
        '',
        tableData.individualAchievementData,
        tableData.individualAchievementData * 5,
      ];
      const newRow12Values = [
        '',
        '',
        'Побед в специальных номинациях',
        '',
        tableData.specialAwardsData,
        tableData.specialAwardsData * 2,
      ];
      const newRow13Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '',
        tableData.lackOfCompetitiveComponentData,
        tableData.lackOfCompetitiveComponentData * 15,
      ];

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
        cellA.alignment = { vertical: 'middle', horizontal: 'center' };

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
        cellC.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };

        [cellA, cellD, cellE, cellF, cellG].forEach((cell) => {
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
        cellG.border = { right: { style: 'thick' }, bottom: { style: 'thin' } };
        cellA.border = { left: { style: 'thick' } };
        cellE.alignment = { vertical: 'middle', horizontal: 'center' };
        cellF.alignment = { vertical: 'middle', horizontal: 'center' };
      }

      const currentDate = new Date();
      const dateString = currentDate.toISOString().slice(0, 10);
      const timeString = currentDate.toTimeString().slice(0, 8).replace(/:/g, '-');
      const fileName = `table_${dateString}_${timeString}.xlsx`;

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
  const arr = ['5.4.1', '5.4.2', '5.4.3', '5.4.4', 5.5];
  const ratio = {
    '5.4.1': [],
    '5.4.2': [5, 3, 2, 0.5],
    '5.4.3': [10, 6, 5, 0.5],
    '5.4.4': [20, 15, 12, 0.5],
    5.5: [30, 20, 18, 0.5],
  };
  try {
    const workbook = new exceljs.Workbook();
    const filePath = 'example.xlsx';

    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet('Лист1');

    if (sheet) {
      console.log('Страница "Лист1" найдена.');
    } else {
      console.log('Страница "Лист1" не найдена.');
    }

    let rowIndex = null;
    let rowEnd = null;
    let rowEndRatio = null;
    let rowEndIndex = null;
    let flagSearch = true;
    let endFlagSearch = true;

    for (let i = 0; i < arr.length; i++) {
      const currentKey = arr[i];
      if (tableData.select === currentKey && ratio.hasOwnProperty(currentKey)) {
        rowEnd = arr[i + 1];
        if (ratio.hasOwnProperty(rowEnd)) {
          rowEndRatio = ratio[rowEnd];
        }
        break;
      }
    }

    console.log(rowEndRatio);

    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData == tableData.select && flagSearch) {
        rowIndex = index;
        console.log(`Значение ${tableData.select} найдено в строке ${index}.`);
        flagSearch = false;
      }
      if (rowData == rowEnd && endFlagSearch) {
        rowEndIndex = index;
        console.log(`Значение ${rowEnd} найдено в строке ${index}.`);
        endFlagSearch = false;
      }
    });

    if (rowIndex !== null) {
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 3;

      const newRow1Values = [
        '',
        tableData.winner,
        'Побед',
        '1-х мест',
        tableData.commandData1,
        tableData.commandData1 * rowEndRatio[0],
      ];
      const newRow2Values = [
        '',
        '',
        'Призовых мест',
        '2-х мест',
        tableData.commandData2,
        tableData.commandData2 * rowEndRatio[1],
      ];
      const newRow3Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '3-х мест',
        tableData.commandData3,
        tableData.commandData3 * rowEndRatio[2],
      ];

      sheet.unMergeCells(`A${rowIndex}:A${rowEndIndex - 1}`);

      sheet.spliceRows(rowIndex + 1, 0, newRow1Values);
      sheet.spliceRows(rowIndex + 2, 0, newRow2Values);
      sheet.spliceRows(rowIndex + 3, 0, newRow3Values);

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
      sheet.unMergeCells(`A${rowEndIndex - 1}:A${rowEndIndex + 2}`);
      sheet.mergeCells(`A${rowIndex}:A${rowEndIndex + 2}`);
      sheet.mergeCells(`B${rowIndex + 1}:B${rowIndex + 3}`);
      sheet.mergeCells(`C${rowIndex + 1}:D${rowIndex + 1}`);
      sheet.mergeCells(`C${rowIndex + 2}:D${rowIndex + 2}`);
      sheet.mergeCells(`C${rowIndex + 3}:D${rowIndex + 3}`);

      for (let i = rowIndex + 1; i <= endRowIndex; i++) {
        const cellA = sheet.getCell(`A${i}`);
        const cellB = sheet.getCell(`B${i}`);
        const cellC = sheet.getCell(`C${i}`);
        const cellD = sheet.getCell(`D${i}`);
        const cellE = sheet.getCell(`E${i}`);
        const cellF = sheet.getCell(`F${i}`);
        const cellG = sheet.getCell(`G${i}`);

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
        cellA.alignment = { vertical: 'middle', horizontal: 'center' };

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
        cellC.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };

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
        cellG.border = { right: { style: 'thick' }, bottom: { style: 'thin' } };
        cellA.border = { left: { style: 'thick' } };
        cellE.alignment = { vertical: 'middle', horizontal: 'center' };
        cellF.alignment = { vertical: 'middle', horizontal: 'center' };
      }

      const currentDate = new Date();
      const dateString = currentDate.toISOString().slice(0, 10);
      const timeString = currentDate.toTimeString().slice(0, 8).replace(/:/g, '-');
      const fileName = `table_${dateString}_${timeString}.xlsx`;

      await workbook.xlsx.writeFile(fileName);

      console.log('Данные успешно вставлены после строки с значением 5.4.');
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 5.4 не найдена.');
    }
  } catch (error) {
    console.error('Ошибка при создании Excel файла:', error);
  }
});

router.post('/createExcel64', async (req, res) => {
  const tableData = req.body;
  const arr = ['6.4.2', '6.4.3', '6.5'];
  const ratio = {
    '6.4.2': [],
    '6.4.3': [5],
    6.5: [50],
  };
  try {
    const workbook = new exceljs.Workbook();
    const filePath = 'example.xlsx';

    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet('Лист1');

    if (sheet) {
      console.log('Страница "Лист1" найдена.');
    } else {
      console.log('Страница "Лист1" не найдена.');
    }

    let rowIndex = null;
    let rowEnd = null;
    let rowEndRatio = null;
    let rowEndIndex = null;
    let flagSearch = true;
    let endFlagSearch = true;

    for (let i = 0; i < arr.length; i++) {
      const currentKey = arr[i];
      if (tableData.select === currentKey && ratio.hasOwnProperty(currentKey)) {
        rowEnd = arr[i + 1];
        if (ratio.hasOwnProperty(rowEnd)) {
          rowEndRatio = ratio[rowEnd];
        }
        break;
      }
    }

    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData == tableData.select && flagSearch) {
        rowIndex = index;
        console.log(`Значение ${tableData.select} найдено в строке ${index}.`);
        flagSearch = false;
      }
      if (rowData == rowEnd && endFlagSearch) {
        rowEndIndex = index;
        console.log(`Значение ${rowEnd} найдено в строке ${index}.`);
        endFlagSearch = false;
      }
    });

    if (rowIndex !== null) {
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 1;

      const newRow1Values = [
        '',
        tableData.winner,
        '',
        '',
        tableData.commandData1,
        tableData.commandData1 * rowEndRatio[0],
      ];

      sheet.unMergeCells(`A${rowIndex}:A${rowEndIndex - 1}`);

      sheet.spliceRows(rowIndex + 1, 0, newRow1Values);

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

      sheet.unMergeCells(`A${rowIndex - 1}:A${rowEndIndex}`);
      sheet.mergeCells(`A${rowIndex}:A${rowEndIndex}`);
      sheet.unMergeCells(`B${rowIndex + 1}:D${rowIndex + 1}`);
      sheet.mergeCells(`B${rowIndex + 1}:D${rowIndex + 1}`);

      for (let i = rowIndex + 1; i <= endRowIndex; i++) {
        const cellA = sheet.getCell(`A${i}`);
        const cellB = sheet.getCell(`B${i}`);
        const cellC = sheet.getCell(`C${i}`);
        const cellD = sheet.getCell(`D${i}`);
        const cellE = sheet.getCell(`E${i}`);
        const cellF = sheet.getCell(`F${i}`);
        const cellG = sheet.getCell(`G${i}`);

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
        cellA.alignment = { vertical: 'middle', horizontal: 'center' };

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
        cellC.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };

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
        cellG.border = { right: { style: 'thick' } };
        cellE.alignment = { vertical: 'middle', horizontal: 'center' };
        cellF.alignment = { vertical: 'middle', horizontal: 'center' };
      }

      const currentDate = new Date();
      const dateString = currentDate.toISOString().slice(0, 10);
      const timeString = currentDate.toTimeString().slice(0, 8).replace(/:/g, '-');
      const fileName = `table_${dateString}_${timeString}.xlsx`;

      await workbook.xlsx.writeFile(fileName);

      console.log('Данные успешно вставлены после строки с значением 6.4.');
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 6.4 не найдена.');
    }
  } catch (error) {
    console.error('Ошибка при создании Excel файла:', error);
  }
});

router.post('/createExcel723', async (req, res) => {
  const tableData = req.body;

  try {
    const workbook = new exceljs.Workbook();
    const filePath = 'example.xlsx';

    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet('Лист1');

    if (sheet) {
      console.log('Страница "Лист1" найдена.');
    } else {
      console.log('Страница "Лист1" не найдена.');
    }

    let rowIndex = null;
    let rowEndIndex = null;
    let flag = true;
    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData == '7.2.3' && flag) {
        rowIndex = index;
        console.log(`Значение 7.2.3 найдено в строке ${index}.`);
        flag = false;
      }
      if (rowData == '7.2.4') {
        rowEndIndex = index;
        console.log(`Значение 7.2.4 найдено в строке ${index}.`);
      }
    });

    if (rowIndex !== null) {
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 7;

      const newRow1Values = [
        '',
        tableData.winner,
        'Призовые места по итогам командного первенства',
        '1-х мест',
        tableData.commandData1,
        tableData.commandData1 * 20,
      ];
      const newRow2Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData2,
        tableData.commandData2 * 15,
      ];
      const newRow3Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData3,
        tableData.commandData3 * 10,
      ];
      const newRow4Values = [
        '',
        '',
        'Призовые места по итогам личного первенства',
        '1-х мест',
        tableData.personalData1,
        tableData.personalData1 * 20,
      ];
      const newRow5Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.personalData2,
        tableData.personalData2 * 15,
      ];
      const newRow6Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.personalData3,
        tableData.personalData3 * 10,
      ];
      const newRow7Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '',
        tableData.lackOfCompetitiveComponentData,
        tableData.lackOfCompetitiveComponentData * 5,
      ];
      sheet.unMergeCells(`A${rowIndex}:A${rowEndIndex - 1}`);
      sheet.spliceRows(rowIndex + 1, 0, newRow1Values);
      sheet.spliceRows(rowIndex + 2, 0, newRow2Values);
      sheet.spliceRows(rowIndex + 3, 0, newRow3Values);
      sheet.spliceRows(rowIndex + 4, 0, newRow4Values);
      sheet.spliceRows(rowIndex + 5, 0, newRow5Values);
      sheet.spliceRows(rowIndex + 6, 0, newRow6Values);
      sheet.spliceRows(rowIndex + 7, 0, newRow7Values);

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
      sheet.unMergeCells(`A${rowIndex}:A${rowEndIndex + 6}`);
      sheet.mergeCells(`A${rowIndex}:A${rowEndIndex + 6}`);
      sheet.mergeCells(`B${rowIndex + 1}:B${rowIndex + 7}`);
      sheet.mergeCells(`C${rowIndex + 1}:C${rowIndex + 3}`);
      sheet.mergeCells(`C${rowIndex + 4}:C${rowIndex + 6}`);
      sheet.mergeCells(`C${rowIndex + 7}:D${rowIndex + 7}`);

      for (let i = rowIndex + 1; i <= endRowIndex; i++) {
        const cellA = sheet.getCell(`A${i}`);
        const cellB = sheet.getCell(`B${i}`);
        const cellC = sheet.getCell(`C${i}`);
        const cellD = sheet.getCell(`D${i}`);
        const cellE = sheet.getCell(`E${i}`);
        const cellF = sheet.getCell(`F${i}`);
        const cellG = sheet.getCell(`G${i}`);

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
        cellA.alignment = { vertical: 'middle', horizontal: 'center' };

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
        cellC.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };

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
        cellG.border = { right: { style: 'thick' }, bottom: { style: 'thin' } };
        cellA.border = { left: { style: 'thick' } };
        cellE.alignment = { vertical: 'middle', horizontal: 'center' };
        cellF.alignment = { vertical: 'middle', horizontal: 'center' };
      }

      const currentDate = new Date();
      const dateString = currentDate.toISOString().slice(0, 10);
      const timeString = currentDate.toTimeString().slice(0, 8).replace(/:/g, '-');
      const fileName = `table_${dateString}_${timeString}.xlsx`;

      await workbook.xlsx.writeFile(fileName);

      console.log('Данные успешно вставлены после строки с значением 7.2.3');
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 7.2.3 не найдена.');
    }
  } catch (error) {
    console.error('Ошибка при создании Excel файла:', error);
  }
});

module.exports = router;
