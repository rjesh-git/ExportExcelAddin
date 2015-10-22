var querystring = require('querystring');
var https = require('https');
var unirest = require('unirest');
var viewID = '', spHostUrl = '', itemsID, listID = '', aToken;

exports.init = function (req, accessToken) {
	viewID = decodeURIComponent(req.query.ViewID);
	listID = decodeURIComponent(req.query.ListID);
	itemsID = decodeURIComponent(req.query.ItemID);
	spHostUrl = decodeURIComponent(req.query.SPHostUrl);
	aToken = accessToken;
};

exports.getViewFields = function () {
	performRequest(spHostUrl + "/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/items");
	//https://xx.sharepoint.com/sites/o365/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/Views(guid'94050996-1781-4820-A95F-5B52CE9D7B0F')/viewFields
};

exports.getListItems = function (viewId) {
};

var performRequest = function (endpoint, success) {
	unirest.get(endpoint)
		.header({ 'Accept': 'application/json', 'Authorization': 'Bearer ' + aToken })
		.send()
		.end(function (response) {
			console.log('endpoint-'+endpoint);
			console.log(response.body);
		});
};