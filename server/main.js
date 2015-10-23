'use strict';
var express = require('express');
var logger = require('connect-logger');
var cookieParser = require('cookie-parser');
var session = require('cookie-session');
var AuthenticationContext = require('adal-node').AuthenticationContext;
var adalHelper = require('./adalHelper.js');
var spHelper = require('./sharepointHelper.js');

var app = express();
app.use(logger());
app.use(cookieParser('a deep secret'));
app.use(session({secret: '1234567890QWERTY'}));

app.get('/report', function(req, res) {
    spHelper.getListItems(req, res);
    
});

app.get('/export', function(req, res) {
  spHelper.init(req);
  adalHelper.processAuth(req, res);
});

app.get('/getAToken', function (req, res) {
  adalHelper.processGetToken(req, res);
});

app.listen(3000);
console.log('listening on 3000');
