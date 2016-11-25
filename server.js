'use strict';

const http = require("http");
const url = require("url");

function start(route, handle) {
  function onRequest(req, res) {
    const pathName = url.parse(req.url).pathname;
    console.log("Request for " + pathName + " received.");

    route(handle, pathName, res, req);
  }

  var port = 1337;

  http.createServer(onRequest).listen(port);
  console.log("Server has started. Listening on port: " + port + "...");
}

exports.start = start;
