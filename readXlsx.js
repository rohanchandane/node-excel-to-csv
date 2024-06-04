const {readSheetNames} = require('read-excel-file/node');

async function getSheetCount(filePath) {
  try {
    const sheet = await readSheetNames(filePath);
    return sheet.length;
  } catch (error) {
    console.error('Error reading Excel file:', error);
    return -1;
  }
}

module.exports = {
  getSheetCount
}