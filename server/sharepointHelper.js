var querystring = require('querystring');
var https = require('https');
var unirest = require('unirest');
var adalHelper = require('./adalHelper.js');
var excelBuilder = require('./excelBuilder.js');

var viewID = '', spHostUrl = '', itemsID, listID = '', resp;

exports.init = function (req) {
	//TODO : Read all below from request form attributes
	viewID = decodeURIComponent(req.query.ViewID);
	listID = decodeURIComponent(req.query.ListID).replace('{', '').replace('}', '');
	itemsID = decodeURIComponent(req.query.ItemID);
	spHostUrl = decodeURIComponent(req.query.SPHostUrl);
	//TODO : Use listname as file name and view name as sheet name
	excelBuilder.init('TheList', 'viewNameHere');
};

exports.getViewFields = function () {
	//TODO: View fields and its type will be available in request form as JSON
	//https://xx.sharepoint.com/sites/o365/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/Views(guid'94050996-1781-4820-A95F-5B52CE9D7B0F')/viewFields
};

exports.getListItems = function (req, res) {
	//TODO: Build REST API from the view fields and selected items
	performRequest(res, spHostUrl + "/_api/web/lists(guid'" + listID + "')/items", processDataItems);
};

var performRequest = function (res, endpoint, success) {
	unirest.get(endpoint)
		.header({ 'Accept': 'application/json', 'Authorization': 'Bearer ' + adalHelper.getAccessToken() })
		.send()
		.end(function (response) {
			excelBuilder.addRows(res, response.body);
		});
};

var processDataItems = function (res, data) {
	//excelBuilder.addRows(res, data);
}