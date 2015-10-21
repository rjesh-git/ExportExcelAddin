var querystring = require('querystring');
var https = require('https');

var viewID = '', spHostUrl = '', itemsID, listID = '', aToken;

exports.init = function (req, accessToken) {
	viewID = decodeURIComponent(req.query.ViewID);
	listID = decodeURIComponent(req.query.ListID);
	itemsID = decodeURIComponent(req.query.ItemID);
	spHostUrl = decodeURIComponent(req.query.SPHostUrl);
	aToken = accessToken;
};

exports.getViewFields = function () {
	console.log('site:'+spHostUrl + ',ListID:'+listID);
	performRequest("/_api/web/lists(guid'81126b95-9b39-4d57-b283-390869cd23e3')", 'GET'
	, function (data) {
		console.log('Fetched ' + data + ' cards');
	});
	//https://xx.sharepoint.com/sites/o365/_api/web/lists(guid'81126B95-9B39-4D57-B283-390869CD23E3')/Views(guid'94050996-1781-4820-A95F-5B52CE9D7B0F')/viewFields
};

exports.getListItems = function (viewId) {
};

var performRequest = function (endpoint, method, success) {
	var options = {
		host: 'tpgbys.sharepoint.com/sites/o365',
		path: endpoint,
		method: method,
		headers: {
			'Accept': 'application/json',
			'Authorization': 'Bearer ' + aToken,
		}
	};
	https.request(options, function (res) {
		res.setEncoding('utf-8');
		var responseString = '';
		res.on('data', function (data) {
			responseString += data;
		});
		res.on('end', function () {
			console.log(responseString);
			var responseObject = JSON.parse(responseString);
			success(responseObject);
		});
	}).end();
};