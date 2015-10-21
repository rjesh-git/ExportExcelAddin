var excelbuilder = require('msexcel-builder');
var workbook,
  sheet1;

//TODO: Use Authcontext and build the query here

//var viewID = GetCurrentCtx().view
//https://xx.sharepoint.com/sites/o365/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/Views(guid'94050996-1781-4820-A95F-5B52CE9D7B0F')/viewFields

exports.init = function (fileName, sheetName) {
  workbook = excelbuilder.createWorkbook('./', fileName + '.xlsx');
  sheet1 = workbook.createSheet(sheetName, 10, 12);
  addTestRows();

  workbook.save(function (ok) {
    if (!ok)
      workbook.cancel();
    else
      console.log('congratulations, your workbook created');
  });
};

var addTestRows = function () {
  sheet1.set(1, 1, 'I am title');
  for (var i = 2; i < 5; i++) {
    sheet1.set(i, 1, 'test' + i);
  }
  workbook.cancel();
};