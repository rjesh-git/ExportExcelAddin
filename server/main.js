'use strict';

var express = require('express');
var logger = require('connect-logger');
var cookieParser = require('cookie-parser');
var session = require('cookie-session');

var AuthenticationContext = require('adal-node').AuthenticationContext;
var adalHelper = require('./adalHelper.js');
var excelBuilder = require('./excelBuilder.js');

var viewID ='', spHostUrl='', itemsID, listID='';

var app = express();
app.use(logger());
app.use(cookieParser('a deep secret'));
app.use(session({secret: '1234567890QWERTY'}));

app.get('/report', function(req, res) {
  //res.header("Content-Type", "text/html");
  excelBuilder.init('TheList','viewNameHere');
  res.download('./' + 'TheList.xlsx');
  //res.send('TODO:Excel will be sent here list: '+ listID +', view:' +viewID);
});

app.get('/export', function(req, res) {
  viewID = decodeURIComponent(req.query.ViewID);
  listID = decodeURIComponent(req.query.ListID);
  itemsID = decodeURIComponent(req.query.ItemID);
  spHostUrl = decodeURIComponent(req.query.SPHostUrl);
  
  adalHelper.processAuth(req, res);
});

app.get('/getAToken', function (req, res) {
  adalHelper.processGetToken(req, res);
});

app.listen(3000);
console.log('listening on 3000');
