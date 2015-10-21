var logger = require('connect-logger');
var cookieParser = require('cookie-parser');
var session = require('cookie-session');
var fs = require('fs');
var crypto = require('crypto');
var AuthenticationContext = require('adal-node').AuthenticationContext;
var accessToken;

process.env['parameters'] = 'parameters.json';

var parametersFile = process.argv[2] || process.env['parameters'];

var sampleParameters;
if (parametersFile) {
  var jsonFile = fs.readFileSync(parametersFile);
  if (jsonFile) {
    sampleParameters = JSON.parse(jsonFile);
  } else {
    console.log('File not found, falling back to defaults: ' + parametersFile);
  }
}

if (!parametersFile) {
  sampleParameters = {
    tenant: 'REPLACETHIS.onmicrosoft.com',
    authorityHostUrl: 'https://login.windows.net',
    clientId: '89bedb2a-3c1e-4df6-b544-8f1f14392ebd',
    clientSecret: 'vM2XycJD8lf29qfckwGC604ATqUBYTFcIsxvdnZuNFo='
  };
}

var authorityUrl = sampleParameters.authorityHostUrl + '/' + sampleParameters.tenant;
var redirectUri = 'http://localhost:3000/getAToken';
var resource = '00000002-0000-0000-c000-000000000000';


var templateAuthzUrl = 'https://login.windows.net/' + sampleParameters.tenant + '/oauth2/authorize?response_type=code&client_id=<client_id>&redirect_uri=<redirect_uri>&state=<state>&resource=<resource>';

function createAuthorizationUrl(state) {
  var authorizationUrl = templateAuthzUrl.replace('<client_id>', sampleParameters.clientId);
  authorizationUrl = authorizationUrl.replace('<redirect_uri>', redirectUri);
  authorizationUrl = authorizationUrl.replace('<state>', state);
  authorizationUrl = authorizationUrl.replace('<resource>', resource);
  return authorizationUrl;
}

exports.processGetToken = function (req, res) {
  if (req.cookies.authstate !== req.query.state) {
    res.send('error: state does not match');
  }
  var authenticationContext = new AuthenticationContext(authorityUrl);
  authenticationContext.acquireTokenWithAuthorizationCode(req.query.code, redirectUri, resource, sampleParameters.clientId, sampleParameters.clientSecret, function (err, response) {
    var message = '';
    if (err) {
      message = 'error: ' + err.message + '\n';
    }
    message += 'response: ' + JSON.stringify(response);

    if (err) {
      res.redirect('report');
      return;
    }

    // TODO: Later, if the access token is expired it can be refreshed.
    authenticationContext.acquireTokenWithRefreshToken(response.refreshToken, sampleParameters.clientId, sampleParameters.clientSecret, resource, function (refreshErr, refreshResponse) {
      if (refreshErr) {
        message += 'refreshError: ' + refreshErr.message + '\n';
      }
      message += 'refreshResponse: ' + JSON.stringify(refreshResponse);
      accessToken = response.accessToken;
      res.redirect('report');
    });
  });
};

exports.processAuth = function (req, res) {
  crypto.randomBytes(48, function (ex, buf) {
    var token = buf.toString('base64').replace(/\//g, '_').replace(/\+/g, '-');
    res.cookie('authstate', token);
    var authorizationUrl = createAuthorizationUrl(token);
    res.redirect(authorizationUrl);
  });
};

exports.getAccessToken = function (){
  return accessToken;
}
