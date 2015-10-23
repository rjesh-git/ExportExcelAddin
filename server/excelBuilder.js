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

exports.addRows = function (data, res) {
  var i = 0, j = 0;
  for (var property in data.value[i]) {
    console.log(property + ':' + data.value[j][property]);
    sheet1.set(j + 1, 1, property); j++;
		}

  for (i = 0; i < data.value.length; i++) {
    j = 0;
    for (var property in data.value[i]) {
      console.log(property + ':' + data.value[i][property]);
      sheet1.set(j + 1, i + 2, data.value[i][property]); j++;
    }
  }
  workbook.save(function (err) {
    if (err)
      throw err;
    else
      //res.download('./' + 'TheList.xlsx');
      console.log('congratulations, your workbook created');
  });
};