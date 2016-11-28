'use strict';

const url = require("url");
const server = require("./server");
const router = require("./router");
const authHelper = require("./authHelper");
const outlook = require("node-outlook");

function getUserEmail(token, callback) {
  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

  const queryParams = {
    '$select': 'DisplayName, EmailAddress'
  };

  outlook.base.getUser({token: token, odataParams: queryParams}, function(error, user){
    if (error) {
      callback(error, null);
    } else {
      callback(null, user.EmailAddress);
    }
  });
}

function tokenReceived(res, error, token) {
  if (error) {
    console.log("Access token error:" + error.message);

    let page = `<p>ERROR: ${error}</p>`;

    res.writeHead(200, {"Content-Type": "text/html"});
    res.write(page);
    res.end();
  } else {
    getUserEmail(token.token.access_token, function(error, email){
      if (error) {
        console.log('getUserEmail returned an error:' + error);

        let page = `<p>ERROR: ${error}</p>`;

        res.writeHead(200, {"Content-Type": "text/html"});
        res.write(page);
        res.end();

      } else if (email) {
        let page = `<p>Email: ${email}</p>
        <p>Access token: ${token.token.access_token}</p>`;

        res.writeHead(200, {"Content-Type": "text/html"});
        res.write(page);
        res.end();
      }
    });
  }
}

let handle = {};

handle["/"] = function (res, req) {
  console.log("Request handler 'home' was called.");

  let homePage = `<p>Please <a href="${authHelper.getAuthUrl()}">sign in</a> with your Office 365 or Outlook.com account.</p>`;

  res.writeHead(200, {"Content-Type": "text/html"});
  res.write(homePage);
  res.end();
};

handle["/authorize"] = function (res, req) {
  console.log("Request handler 'authorize' was called.");

  let urlParts = url.parse(req.url, true);
  let code = urlParts.query.code;

  console.log("Code: " + code);

  authHelper.getTokenFromCode(code, tokenReceived, res);
};

server.start(router.route, handle);
