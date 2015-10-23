var querystring = require('querystring');
var https = require('https');
var unirest = require('unirest');
var adalHelper = require('./adalHelper.js');
var excelBuilder = require('./excelBuilder.js');

var viewID = '', spHostUrl = '', itemsID, listID = '', resp;

exports.init = function (req) {
	viewID = decodeURIComponent(req.query.ViewID);
	listID = decodeURIComponent(req.query.ListID);
	itemsID = decodeURIComponent(req.query.ItemID);
	spHostUrl = decodeURIComponent(req.query.SPHostUrl);
	excelBuilder.init('TheList','viewNameHere');
};

exports.getViewFields = function () {
	//https://xx.sharepoint.com/sites/o365/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/Views(guid'94050996-1781-4820-A95F-5B52CE9D7B0F')/viewFields
};

exports.getListItems = function (complete) {
	//excelBuilder.addRows(performRequest(spHostUrl + "/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/items", processDataItems));
	performRequest(spHostUrl + "/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/items", processDataItems);
};

var performRequest = function (endpoint, success) {
	unirest.get(endpoint)
		.header({ 'Accept': 'application/json', 'Authorization': 'Bearer ' + adalHelper.getAccessToken() })
		.send()
		.end(function (response) {
			success(response.body);
		});
};

var processDataItems = function (data) {
	/*
	for (var i = 0; i < data.value.length; i++) {
		for (var property in data.value[i]) {
			console.log(property +':'+ data.value[i][property]);
		}
	}*/
	excelBuilder.addRows(data);
}