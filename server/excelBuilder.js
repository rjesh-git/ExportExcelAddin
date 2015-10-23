var excelbuilder = require('msexcel-builder');
var workbook,
  sheet1;

//TODO: Use Authcontext and build the query here

exports.init = function (fileName, sheetName) {
  workbook = excelbuilder.createWorkbook('./', fileName + '.xlsx');
  sheet1 = workbook.createSheet(sheetName, 500, 500);
};

var addTestRows = function () {
  sheet1.set(1, 1, 'I am title');
  for (var i = 2; i < 5; i++) {
    sheet1.set(i, 1, 'test' + i);
  }
  workbook.cancel();
};

exports.addRows = function (data) {
  var i, j;
  for (j = 0; j < data.value[0].length; j++) {
    sheet1.set(1, j, data.value[j]);
  }
  for (i = 0; i < data.value.length; i++) {
    for (j = 0; j < data.value[i].length; j++) {
      sheet1.set(i, j, data.value[j]);
    }
  }
  workbook.save(function (ok) {
    if (!ok)
      workbook.cancel();
    else
      console.log('congratulations, your workbook created');
  });
};