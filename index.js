'use strict';

const url = require("url");
const server = require("./server");
const router = require("./router");
const authHelper = require("./authHelper");

const homePage = `
<p>Please <a href="${authHelper.getAuthUrl()}">sign in</a> with your Office 365 or Outlook.com account.</p>'`;

let handle = {};

handle["/"] = function (res, req) {
  console.log("Request handler 'home' was called.");

  res.writeHead(200, {"Content-Type": "text/html"});
  res.write(homePage);
  res.end();
};

handle["/authorize"] = function (res, req) {
  console.log("Request handler 'authorize' was called.");

  debugger;
  let urlParts = url.parse(req.url, true);
  let code = urlParts.query.code;

  console.log("Code: " + code);

  let authorizePage = `<p>Receive auth code: ${code}</p>`;
  res.writeHead(200, {"Content-Type": "text/html"});
  res.write(authorizePage);
  res.end();
};

server.start(router.route, handle);
