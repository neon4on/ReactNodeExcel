// Модули
const express = require('express');
const router = express.Router();
const fs = require('fs');
const path = require('path');
const exceljs = require('exceljs');
const axios = require('axios');
const lockfile = require('lockfile');

// Константы
const NUM_PARAM = '1986228881';
const newFolderPath = 'new';
const oldFolderPath = 'old';

// Функция для проверки открытости файла
const isFileLocked = (filePath) => {
  try {
    fs.renameSync(filePath, filePath);
    return false;
  } catch (error) {
    if (error.code === 'EBUSY') {
      return true;
    }
    throw error;
  }
};

const checkLocksInFolder = () => {
  const files = fs.readdirSync(newFolderPath);
  files.forEach((file) => {
    const filePath = path.join(newFolderPath, file);
    try {
      fs.renameSync(filePath, filePath);
      console.log(`Файл ${filePath} не заблокирован.`);
    } catch (error) {
      if (error.code === 'EBUSY') {
        console.log(`Файл ${filePath} заблокирован другим процессом.`);
      } else {
        console.error(`Ошибка при проверке файла ${filePath}:`, error);
      }
    }
  });
};

// Форма под пунктом 5.1
router.post('/createExcel51', async function (req, res) {
  const tableData = req.body;

  try {
    const workbook = new exceljs.Workbook();

    if (!fs.existsSync(newFolderPath)) {
      fs.mkdirSync(newFolderPath);
    }
    if (!fs.existsSync(oldFolderPath)) {
      fs.mkdirSync(oldFolderPath);
    }

    const files = fs.readdirSync(newFolderPath);

    let filePath = null;
    for (const file of files) {
      if (file.startsWith('table')) {
        filePath = path.join(newFolderPath, file);
        break;
      }
    }

    if (!filePath) {
      throw new Error('Файл с префиксом "table" не найден в папке "new"');
    }

    if (isFileLocked(filePath)) {
      console.log('Файл заблокирован другим процессом.');
      await axios.get(
        `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=файл%20открыт.%20данные%20не%20добавлены`,
      );
      return;
    }

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
      if (rowData == 5.1) {
        rowIndex = index;
        console.log(`Значение 5.1 найдено в строке ${index}.`);
      }
      if (rowData == 5.2) {
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
      console.log(tableData);
      const newRow1Values = [
        '5.1.1',
        tableData.winner,
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData1 ? tableData.commandData1 : '',
        tableData.commandData1 ? tableData.commandData1 * 30 : '',
      ];
      const newRow2Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData2 ? tableData.commandData2 : '',
        tableData.commandData2 ? tableData.commandData2 * 25 : '',
      ];
      const newRow3Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData3 ? tableData.commandData3 : '',
        tableData.commandData3 ? tableData.commandData3 * 20 : '',
      ];
      const newRow4Values = [
        '',
        '',
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData11 ? tableData.commandData11 : '',
        tableData.commandData11 ? tableData.commandData11 * 10 : '',
      ];
      const newRow5Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData21 ? tableData.commandData21 : '',
        tableData.commandData21 ? tableData.commandData21 * 8 : '',
      ];
      const newRow6Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData31 ? tableData.commandData31 : '',
        tableData.commandData31 ? tableData.commandData31 * 6 : '',
      ];
      const newRow7Values = [
        '',
        '',
        'Призовые места по итогам личного первенства в номинациях',
        '1-х мест',
        tableData.personalData1 ? tableData.personalData1 : '',
        tableData.personalData1 ? tableData.personalData1 * 5 : '',
      ];
      const newRow8Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.personalData2 ? tableData.personalData2 : '',
        tableData.personalData2 ? tableData.personalData2 * 4 : '',
      ];
      const newRow9Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.personalData3 ? tableData.personalData3 : '',
        tableData.personalData3 ? tableData.personalData3 * 3 : '',
      ];
      const newRow10Values = [
        '',
        '',
        'Гран При',
        '',
        tableData.grandPrizeData ? tableData.grandPrizeData : '',
        tableData.grandPrizeData ? tableData.grandPrizeData * 15 : '',
      ];
      const newRow11Values = [
        '',
        '',
        'Приз за отдельные достижения',
        '',
        tableData.individualAchievementData ? tableData.individualAchievementData : '',
        tableData.individualAchievementData ? tableData.individualAchievementData * 5 : '',
      ];
      const newRow12Values = [
        '',
        '',
        'Побед в специальных номинациях',
        '',
        tableData.specialAwardsData ? tableData.specialAwardsData : '',
        tableData.specialAwardsData ? tableData.specialAwardsData * 2 : '',
      ];
      const newRow13Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '',
        tableData.lackOfCompetitiveComponentData ? tableData.lackOfCompetitiveComponentData : '',
        tableData.lackOfCompetitiveComponentData
          ? tableData.lackOfCompetitiveComponentData * 15
          : '',
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
      console.log(tableData);
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

      const newFilePath = path.join(newFolderPath, fileName);
      fs.renameSync(filePath, path.join(oldFolderPath, path.basename(filePath)));
      await workbook.xlsx.writeFile(newFilePath);

      await axios.get(`http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=данные%20были%20добавлены`);
      console.log('Данные успешно вставлены после строки с значением 5.1.');
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 5.1 не найдена.');
    }
  } catch (error) {
    await axios.get(
      `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=ошибка%20при%20открытии%20файла`,
    );
    console.error('Ошибка при создании Excel файла:', error);
  }
});

/// Форма под пунктом 5.2
router.post('/createExcel52', async function (req, res) {
  const tableData = req.body;

  try {
    const workbook = new exceljs.Workbook();

    if (!fs.existsSync(newFolderPath)) {
      fs.mkdirSync(newFolderPath);
    }
    if (!fs.existsSync(oldFolderPath)) {
      fs.mkdirSync(oldFolderPath);
    }

    const files = fs.readdirSync(newFolderPath);

    let filePath = null;
    for (const file of files) {
      if (file.startsWith('table')) {
        filePath = path.join(newFolderPath, file);
        break;
      }
    }

    if (!filePath) {
      throw new Error('Файл с префиксом "table" не найден в папке "new"');
    }

    if (isFileLocked(filePath)) {
      console.log('Файл заблокирован другим процессом.');
      await axios.get(
        `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=файл%20открыт.%20данные%20не%20добавлены`,
      );
      return;
    }

    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet('Лист1');

    if (sheet) {
      console.log('Страница "Лист1" найдена.');
    } else {
      console.log('Страница "Лист1" не найдена.');
    }

    let rowIndex = null;
    let rowEndIndex = null;
    let flagSearch = true;

    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData && rowData == 5.2) {
        rowIndex = index;
        console.log(`Значение 5.2 найдено в строке ${index}.`);
      }
      if (rowData == 5.3 && flagSearch) {
        rowEndIndex = index;
        console.log(`Значение 5.3 найдено в строке ${index}.`);
        flagSearch = false;
      }
    });
    console.log(tableData);
    if (rowIndex !== null) {
      console.log(`Строка с 5.2 найдена: ${rowIndex}`);

      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 13;

      let j = 0;
      for (let i = startRowIndex + 1; i < rowEndIndex; i++) {
        if (j % 13 == 0) {
          const cellValue = sheet.getCell(`A${i}`).value;
          if (cellValue !== null) {
            const parts = cellValue.split('.');
            const number = parseInt(parts[2]);
            parts[2] = (number + 1).toString();
            sheet.getCell(`A${i}`).value = parts.join('.');
          }
        }
        j++;
      }

      const newRow1Values = [
        '5.2.1',
        tableData.winner,
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData1 ? tableData.commandData1 : '',
        tableData.commandData1 ? tableData.commandData1 * 30 : '',
      ];
      const newRow2Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData2 ? tableData.commandData2 : '',
        tableData.commandData2 ? tableData.commandData2 * 25 : '',
      ];
      const newRow3Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData3 ? tableData.commandData3 : '',
        tableData.commandData3 ? tableData.commandData3 * 20 : '',
      ];
      const newRow4Values = [
        '',
        '',
        'Призовые места по итогам командного первенства в номинациях',
        '1-х мест',
        tableData.commandData11 ? tableData.commandData11 : '',
        tableData.commandData11 ? tableData.commandData11 * 10 : '',
      ];
      const newRow5Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData21 ? tableData.commandData21 : '',
        tableData.commandData21 ? tableData.commandData21 * 8 : '',
      ];
      const newRow6Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData31 ? tableData.commandData31 : '',
        tableData.commandData31 ? tableData.commandData31 * 6 : '',
      ];
      const newRow7Values = [
        '',
        '',
        'Призовые места по итогам личного первенства в номинациях',
        '1-х мест',
        tableData.personalData1 ? tableData.personalData1 : '',
        tableData.personalData1 ? tableData.personalData1 * 5 : '',
      ];
      const newRow8Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.personalData2 ? tableData.personalData2 : '',
        tableData.personalData2 ? tableData.personalData2 * 4 : '',
      ];
      const newRow9Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.personalData3 ? tableData.personalData3 : '',
        tableData.personalData3 ? tableData.personalData3 * 3 : '',
      ];
      const newRow10Values = [
        '',
        '',
        'Гран При',
        '',
        tableData.grandPrizeData ? tableData.grandPrizeData : '',
        tableData.grandPrizeData ? tableData.grandPrizeData * 15 : '',
      ];
      const newRow11Values = [
        '',
        '',
        'Приз за отдельные достижения',
        '',
        tableData.individualAchievementData ? tableData.individualAchievementData : '',
        tableData.individualAchievementData ? tableData.individualAchievementData * 5 : '',
      ];
      const newRow12Values = [
        '',
        '',
        'Побед в специальных номинациях',
        '',
        tableData.specialAwardsData ? tableData.specialAwardsData : '',
        tableData.specialAwardsData ? tableData.specialAwardsData * 2 : '',
      ];
      const newRow13Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '',
        tableData.lackOfCompetitiveComponentData ? tableData.lackOfCompetitiveComponentData : '',
        tableData.lackOfCompetitiveComponentData
          ? tableData.lackOfCompetitiveComponentData * 15
          : '',
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

      // Сначала снимаем объединение ячеек внутри внешнего цикла
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
      console.log(tableData);
      sheet.mergeCells(`A${startRowIndex}:A${endRowIndex}`);
      sheet.mergeCells(`B${startRowIndex}:B${endRowIndex}`);
      sheet.mergeCells(`C${startRowIndex}:C${startRowIndex + 2}`);
      sheet.mergeCells(`C${startRowIndex + 3}:C${startRowIndex + 5}`);
      sheet.mergeCells(`C${startRowIndex + 6}:C${startRowIndex + 8}`);
      sheet.mergeCells(`C${startRowIndex + 9}:D${startRowIndex + 9}`);
      sheet.mergeCells(`C${startRowIndex + 10}:D${startRowIndex + 10}`);
      sheet.mergeCells(`C${startRowIndex + 11}:D${startRowIndex + 11}`);
      sheet.mergeCells(`C${startRowIndex + 12}:D${startRowIndex + 12}`);

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

      const newFilePath = path.join(newFolderPath, fileName);

      fs.renameSync(filePath, path.join(oldFolderPath, path.basename(filePath)));
      await workbook.xlsx.writeFile(newFilePath);
      await axios.get(`http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=данные%20были%20добавлены`);
      console.log('Данные успешно вставлены после строки с значением 5.2.');
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 5.2 не найдена.');
    }
  } catch (error) {
    await axios.get(
      `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=ошибка%20при%20открытии%20файла`,
    );
    console.error('Ошибка при создании Excel файла:', error);
  }
});

// Форма под пунктом 5.4
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

    if (!fs.existsSync(newFolderPath)) {
      fs.mkdirSync(newFolderPath);
    }
    if (!fs.existsSync(oldFolderPath)) {
      fs.mkdirSync(oldFolderPath);
    }

    const files = fs.readdirSync(newFolderPath);

    let filePath = null;
    for (const file of files) {
      if (file.startsWith('table')) {
        filePath = path.join(newFolderPath, file);
        break;
      }
    }

    if (!filePath) {
      throw new Error('Файл с префиксом "table" не найден в папке "new"');
    }

    if (isFileLocked(filePath)) {
      console.log('Файл заблокирован другим процессом.');
      await axios.get(
        `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=файл%20открыт.%20данные%20не%20добавлены`,
      );
      return;
    }

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
    console.log(tableData);
    if (rowIndex !== null) {
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 3;

      const newRow1Values = [
        '',
        tableData.winner,
        'Побед',
        '1-х мест',
        tableData.commandData1 ? tableData.commandData1 : '',
        tableData.commandData1 ? tableData.commandData1 * rowEndRatio[0] : '',
      ];
      const newRow2Values = [
        '',
        '',
        'Призовых мест',
        '2-х мест',
        tableData.commandData2 ? tableData.commandData2 : '',
        tableData.commandData2 ? tableData.commandData2 * rowEndRatio[1] : '',
      ];
      const newRow3Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '3-х мест',
        tableData.commandData3 ? tableData.commandData3 : '',
        tableData.commandData3 ? tableData.commandData3 * rowEndRatio[2] : '',
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

      const newFilePath = path.join(newFolderPath, fileName);

      fs.renameSync(filePath, path.join(oldFolderPath, path.basename(filePath)));
      await workbook.xlsx.writeFile(newFilePath);
      console.log('Данные успешно вставлены после строки с значением 5.4.');
      await axios.get(`http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=данные%20были%20добавлены`);
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 5.4 не найдена.');
    }
  } catch (error) {
    await axios.get(
      `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=ошибка%20при%20открытии%20файла`,
    );
    console.error('Ошибка при создании Excel файла:', error);
  }
});

// Форма под пунктом 6.4
router.post('/createExcel64', async (req, res) => {
  const tableData = req.body;
  const arr = ['6.4.1', '6.4.2', '6.4.3', '6.4.4', '6.5'];
  const ratio = {
    '6.4.1': [],
    '6.4.2': [1],
    '6.4.3': [5],
    '6.4.4': [50],
    6.5: [100],
  };
  try {
    const workbook = new exceljs.Workbook();

    if (!fs.existsSync(newFolderPath)) {
      fs.mkdirSync(newFolderPath);
    }
    if (!fs.existsSync(oldFolderPath)) {
      fs.mkdirSync(oldFolderPath);
    }

    const files = fs.readdirSync(newFolderPath);

    let filePath = null;
    for (const file of files) {
      if (file.startsWith('table')) {
        filePath = path.join(newFolderPath, file);
        break;
      }
    }

    if (!filePath) {
      throw new Error('Файл с префиксом "table" не найден в папке "new"');
    }

    if (isFileLocked(filePath)) {
      console.log('Файл заблокирован другим процессом.');
      await axios.get(
        `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=файл%20открыт.%20данные%20не%20добавлены`,
      );
      return;
    }

    console.log('ОТКРЫВАЕМ ФАЙЛ!');

    try {
      await workbook.xlsx.readFile(filePath);
    } catch (error) {
      console.error('Ошибка при чтении файла Excel:', error);
    }

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

    console.log('НАЧИНАЕМ!');

    let rowIndex64_1 = null;
    let rowIndex64_2 = null;
    let rowIndex64_3 = null;
    let rowIndex64_4 = null;
    let rowIndex64_1_Flag = true;
    let rowIndex64_2_Flag = true;
    let rowIndex64_3_Flag = true;
    let rowIndex64_4_Flag = true;
    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData === '6.4.1' && rowIndex64_1_Flag) {
        rowIndex64_1 = index;
        rowIndex64_1_Flag = false;
      }
      if (rowData === '6.4.4' && rowIndex64_4_Flag) {
        rowIndex64_4 = index;
        rowIndex64_4_Flag = false;
      }
      if (rowData === '6.4.2' && rowIndex64_2_Flag) {
        rowIndex64_2 = index;
        rowIndex64_2_Flag = false;
      }
      if ((rowData === '6.5' || rowData === 6.5) && rowIndex64_3_Flag) {
        rowIndex64_3 = index;
        rowIndex64_3_Flag = false;
      }
    });

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
    console.log(tableData);

    if (!rowIndex64_1) {
      const newRow = ['6.4.1', 'Муниципальный Уровень', '', '', '', ''];
      sheet.spliceRows(rowIndex64_2, 0, newRow);
      sheet.unMergeCells(`B${rowIndex64_2}:G${rowIndex64_2}`);
      sheet.mergeCells(`B${rowIndex64_2}:G${rowIndex64_2}`);
      const mergedCell = sheet.getCell(`B${rowIndex64_2}`);
      mergedCell.style = {
        font: { bold: true, name: 'Times New Roman', size: 12 },
        alignment: { vertical: 'middle', horizontal: 'left', wrapText: true },
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thick' },
        },
      };
    }

    if (!rowIndex64_4) {
      const newRow = ['6.4.4', 'Международный Уровень', '', '', '', ''];
      sheet.spliceRows(rowIndex64_3 + 1, 0, newRow);
      sheet.unMergeCells(`B${rowIndex64_3 + 1}:G${rowIndex64_3 + 1}`);
      sheet.mergeCells(`B${rowIndex64_3 + 1}:G${rowIndex64_3 + 1}`);
      const mergedCell = sheet.getCell(`B${rowIndex64_3 + 1}`);
      mergedCell.style = {
        font: { bold: true, name: 'Times New Roman', size: 12 },
        alignment: { vertical: 'middle', horizontal: 'left', wrapText: true },
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thick' },
        },
      };
    }

    console.log('ПРИВЕТ!');

    sheet.eachRow({ includeEmpty: false }, (row, index) => {
      const rowData = row.getCell(1).value;
      if (rowData == tableData.select && flagSearch) {
        rowIndex = index;
        console.log(`1 Значение ${tableData.select} найдено в строке ${index}.`);
        flagSearch = false;
      }
      if (rowData == rowEnd && endFlagSearch) {
        rowEndIndex = index;
        console.log(`2 Значение ${rowEnd} найдено в строке ${index}.`);
        endFlagSearch = false;
      }
    });

    console.log('НАШЁЛ!');

    if (rowIndex !== null) {
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 1;

      const newRow1Values = [
        '',
        tableData.winner,
        '',
        '',
        tableData.commandData1 ? tableData.commandData1 : '',
        tableData.commandData1 ? tableData.commandData1 * rowEndRatio[0] : '',
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

      sheet.unMergeCells(`A${rowIndex}:A${rowEndIndex}`);
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

      let cellA1 = 4;
      let cellA2 = 4;
      let cellA3 = 4;
      let cellA4 = 4;
      let cellA1Flag = true;
      let cellA2Flag = true;
      let cellA3Flag = true;
      let cellA4Flag = true;
      sheet.eachRow({ includeEmpty: false }, (row, index) => {
        const rowData = row.getCell(1).value;
        if (rowData === '6.4.1' && cellA1Flag) {
          const startCell = sheet.getCell(`B${index}`);
          const endCell = sheet.getCell(`G${index}`);
          applyStyle(startCell, endCell);
          cellA1Flag = false;
        }
        if (rowData === '6.4.2' && cellA2Flag) {
          const startCell = sheet.getCell(`B${index}`);
          const endCell = sheet.getCell(`G${index}`);
          applyStyle(startCell, endCell);
          cellA2Flag = false;
        }
        if (rowData === '6.4.3' && cellA3Flag) {
          const startCell = sheet.getCell(`B${index}`);
          const endCell = sheet.getCell(`G${index}`);
          applyStyle(startCell, endCell);
          cellA3Flag = false;
        }
        if (rowData === '6.4.4' && cellA4Flag) {
          const startCell = sheet.getCell(`B${index}`);
          const endCell = sheet.getCell(`G${index}`);
          applyStyle(startCell, endCell);
          cellA4Flag = false;
        }
      });

      function applyStyle(startCell, endCell) {
        for (let i = startCell.row; i <= endCell.row; i++) {
          for (let j = startCell.col; j <= endCell.col; j++) {
            const cell = sheet.getCell(i, j);
            cell.style = {
              font: { bold: true, name: 'Times New Roman', size: 12 },
              alignment: { vertical: 'middle', horizontal: 'left', wrapText: true },
              border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thick' },
              },
            };
          }
        }
      }

      const currentDate = new Date();
      const dateString = currentDate.toISOString().slice(0, 10);
      const timeString = currentDate.toTimeString().slice(0, 8).replace(/:/g, '-');
      const fileName = `table_${dateString}_${timeString}.xlsx`;

      const newFilePath = path.join(newFolderPath, fileName);

      fs.renameSync(filePath, path.join(oldFolderPath, path.basename(filePath)));
      await workbook.xlsx.writeFile(newFilePath);

      console.log('Данные успешно вставлены после строки с значением 6.4.');
      await axios.get(`http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=данные%20были%20добавлены`);
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 6.4 не найдена.');
    }
  } catch (error) {
    await axios.get(
      `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=ошибка%20при%20открытии%20файла`,
    );
    console.error('Ошибка при создании Excel файла:', error);
  }
});

// Форма под пунктом 7.2.3
router.post('/createExcel723', async (req, res) => {
  const tableData = req.body;

  try {
    const workbook = new exceljs.Workbook();

    if (!fs.existsSync(newFolderPath)) {
      fs.mkdirSync(newFolderPath);
    }
    if (!fs.existsSync(oldFolderPath)) {
      fs.mkdirSync(oldFolderPath);
    }

    const files = fs.readdirSync(newFolderPath);

    let filePath = null;
    for (const file of files) {
      if (file.startsWith('table')) {
        filePath = path.join(newFolderPath, file);
        break;
      }
    }

    if (!filePath) {
      throw new Error('Файл с префиксом "table" не найден в папке "new"');
    }

    if (isFileLocked(filePath)) {
      console.log('Файл заблокирован другим процессом.');
      await axios.get(
        `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=файл%20открыт.%20данные%20не%20добавлены`,
      );
      return;
    }

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
    console.log(tableData);
    if (rowIndex !== null) {
      const startRowIndex = rowIndex + 1;
      const endRowIndex = rowIndex + 7;

      const newRow1Values = [
        '',
        tableData.winner,
        'Призовые места по итогам командного первенства',
        '1-х мест',
        tableData.commandData1 ? tableData.commandData1 : '',
        tableData.commandData1 ? tableData.commandData1 * 20 : '',
      ];
      const newRow2Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.commandData2 ? tableData.commandData2 : '',
        tableData.commandData2 ? tableData.commandData2 * 15 : '',
      ];
      const newRow3Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.commandData3 ? tableData.commandData3 : '',
        tableData.commandData3 ? tableData.commandData3 * 10 : '',
      ];
      const newRow4Values = [
        '',
        '',
        'Призовые места по итогам личного первенства',
        '1-х мест',
        tableData.personalData1 ? tableData.personalData1 : '',
        tableData.personalData1 ? tableData.personalData1 * 20 : '',
      ];
      const newRow5Values = [
        '',
        '',
        '',
        '2-х мест',
        tableData.personalData2 ? tableData.personalData2 : '',
        tableData.personalData2 ? tableData.personalData2 * 15 : '',
      ];
      const newRow6Values = [
        '',
        '',
        '',
        '3-х мест',
        tableData.personalData3 ? tableData.personalData3 : '',
        tableData.personalData3 ? tableData.personalData3 * 10 : '',
      ];
      const newRow7Values = [
        '',
        '',
        'Отсутствие соревновательной составляющей',
        '',
        tableData.lackOfCompetitiveComponentData ? tableData.lackOfCompetitiveComponentData : '',
        tableData.lackOfCompetitiveComponentData
          ? tableData.lackOfCompetitiveComponentData * 5
          : '',
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

      const newFilePath = path.join(newFolderPath, fileName);

      fs.renameSync(filePath, path.join(oldFolderPath, path.basename(filePath)));
      await workbook.xlsx.writeFile(newFilePath);

      console.log('Данные успешно вставлены после строки с значением 7.2.3');
      await axios.get(`http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=данные%20были%20добавлены`);
      res.send('Данные успешно вставлены');
    } else {
      console.log('Строка с 7.2.3 не найдена.');
    }
  } catch (error) {
    await axios.get(
      `http://home.teyhd.ru:3334/?num=${NUM_PARAM}&msg=ошибка%20при%20открытии%20файла`,
    );
    console.error('Ошибка при создании Excel файла:', error);
  }
});

module.exports = router;
