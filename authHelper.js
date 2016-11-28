'use strict';

const local = require("./local.js");

const credentials = {
  client: {
    id: local.outlook.appId,
    secret: local.outlook.password
  },
  auth: {
    tokenHost: "https://login.microsoftonline.com", 
    authorizePath: "/common/oauth2/v2.0/authorize",
    tokenPath: "/common/oauth2/v2.0/token"
  }
}

const oauth2 = require("simple-oauth2").create(credentials);

const redirectUri = "http://localhost:1337/authorize";

const scopes = [
  "openid",
  "https://outlook.office.com/mail.read"
];

function getAuthUrl() {
  const returnVal = oauth2.authorizationCode.authorizeURL({
    redirect_uri: redirectUri,
    scope: scopes.join(" ")
  });

  console.log("Generated auth url: " + returnVal);
  return returnVal;
}

function getTokenFromCode(authCode, callback, response) {
  let token;
  oauth2.authorizationCode.getToken({
    code: authCode,
    redirect_uri: redirectUri,
    scope: scopes.join(" ")
  }, function (error, result){
    if (error) {
      console.log("Access token error: ", error.message);
      callback(response, error, null);
    } else {
      token = oauth2.accessToken.create(result);
      console.log("Token created: ", token.token);
      callback(response, null, token);
    }
  });
}

exports.getAuthUrl = getAuthUrl;
exports.getTokenFromCode = getTokenFromCode;
