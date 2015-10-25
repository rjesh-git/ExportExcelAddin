var excelbuilder = require('msexcel-builder');
var workbook, fileName,
  sheet1;

exports.init = function (listName, sheetName) {
  fileName = listName + '.xlsx';
  workbook = excelbuilder.createWorkbook('./', fileName);
  sheet1 = workbook.createSheet(sheetName, 500, 500);
};
exports.addRows = function (res, data, listFields) {
  var i = 0, j, pValue;
  for (j = 0; j < listFields.length; j++) {
    sheet1.set(j + 1, 1, listFields[j].DisplayName);
  }

  for (i = 0; i < data.value.length; i++) {
    for (j = 0; j < listFields.length; j++) {
      switch (listFields[j].FieldType) {
        case 'User':
        case 'Lookup':
          pValue = data.value[i][listFields[j].RealFieldName].Title;
          break;
        case 'TaxonomyFieldType':
          pValue = data.value[i][listFields[j].RealFieldName].TermGuid;
          break;
        default:
          pValue = data.value[i][listFields[j].RealFieldName];
          break;
      }
      sheet1.set(j + 1, i + 2, pValue);
    }
  }

  workbook.save(function (err) {
    if (err)
      throw err;
    else {
      console.log('congratulations, your workbook created');
      res.download('./' + fileName);
    }
  });
};