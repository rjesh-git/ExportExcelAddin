var unirest = require('unirest');
var adalHelper = require('./adalHelper.js');
var excelBuilder = require('./excelBuilder.js');
var viewID = '', spHostUrl = '', itemIDs, listID = '', listFields, listName;

exports.init = function (req) {
	listFields = JSON.parse(req.body.ListFields);
	viewID = req.body.ViewID;
	listID = req.body.ListID;
	listName = req.body.ListName;
	itemIDs = req.body.ItemIDs.split(',');
	spHostUrl = req.body.WebUrl;
	//TODO : Use view name as sheet name
	excelBuilder.init(listName, 'viewNameHere');
};

exports.getListItems = function (req, res) {
	performRequest(res, buildQuery(), processDataItems);
};

var performRequest = function (res, endpoint, success) {
	unirest.get(endpoint)
		.header({ 'Accept': 'application/json', 'Authorization': 'Bearer ' + adalHelper.getAccessToken() })
		.send()
		.end(function (response) {
			excelBuilder.addRows(res, response.body, listFields);
		});
};

var processDataItems = function (res, data) {
	//excelBuilder.addRows(res, data);
};

var buildQuery = function () {
	var filter = '$filter=(', select = '$select=', expand = '$expand=', exFlag,
		query = spHostUrl + "/_api/web/lists(guid'" + listID + "')/items?";
	for (var i = 0; i < itemIDs.length; i++) {
		filter += '(ID eq ' + itemIDs[i] + ')';
		filter += (i !== itemIDs.length - 1) ? ' or ' : ')';
	}
	for (i = 0; i < listFields.length; i++) {
		exFlag = false;
		switch (listFields[i].FieldType) {
			case 'User':
				select += listFields[i].RealFieldName + '/Title';
				expand += listFields[i].RealFieldName + '/ID';
				exFlag = true;
				break;
			case 'Lookup':
				select += listFields[i].RealFieldName + '/Title';
				expand += listFields[i].RealFieldName + '/Title';
				exFlag = true;
				break;
			default:
				select += listFields[i].RealFieldName;
				break;
		}
		if (i !== listFields.length - 1) {
			select += ',';
			expand += exFlag ? ',' : '';
		}
	}
	return (query + filter + '&' + select + '&' + expand);
};