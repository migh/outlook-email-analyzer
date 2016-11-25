'use strict';

const server = require("./server");
const router = require("./router");

let handle = {};

function home(res, req) {
  console.log("Request handler 'home' was called.");

  res.writeHead(200, {"Content-Type": "text/html"});
  res.write('<p>Hello world!</p>');
  res.end();
}

handle["/"] = home;

server.start(router.route, handle);
