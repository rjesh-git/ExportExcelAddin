var express = require('express');
var app = express();

var fname = Math.random() + '.xlsx';
var excelbuilder = require('msexcel-builder');
var workbook = excelbuilder.createWorkbook('./', fname)
var sheet1 = workbook.createSheet('sheet1', 10, 12);

var buffer;

sheet1.set(1, 1, 'I am title');
for (var i = 2; i < 5; i++)
sheet1.set(i, 1, 'test' + i);

/*workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        buffer = zip.generate({type: "nodebuffer"});
      }
  });*/
   
// Save it 
workbook.save(function(ok){
if (!ok) 
  workbook.cancel();
else
  console.log('congratulations, your workbook created');
});
 
app.get('/', function (req, res) {
  res.download('./' + fname);
});

var server = app.listen(3000, function () {
  
});


//var viewID = GetCurrentCtx().view
//https://xx.sharepoint.com/sites/o365/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/Views(guid'94050996-1781-4820-A95F-5B52CE9D7B0F')/viewFields