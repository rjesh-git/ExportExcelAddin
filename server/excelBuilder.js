var excelbuilder = require('msexcel-builder');
var workbook,fileName,
  sheet1;

exports.init = function (listName, sheetName) {
  fileName = listName + '.xlsx';
  workbook = excelbuilder.createWorkbook('./', fileName);
  sheet1 = workbook.createSheet(sheetName, 500, 500);
};
exports.addRows = function (res, data) {
  var i = 0, j = 0;
  for (var property in data.value[i]) {
    sheet1.set(j + 1, 1, property); j++;
		}
  for (i = 0; i < data.value.length; i++) {
    j = 0;
    for (var property in data.value[i]) {
      sheet1.set(j + 1, i + 2, data.value[i][property]); j++;
    }
  }
  workbook.save(function (err) {
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
      res.download('./' + fileName);
  });
};