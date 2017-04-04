'use strict';

const url = require("url");
const server = require("./server");
const router = require("./router");
const authHelper = require("./authHelper");
const outlook = require("node-outlook");

function getValueFromCookie(valueName, cookie){
  if (cookie.indexOf(valueName) !== -1) {
    let start = cookie.indexOf(valueName) + valueName.length + 1;
    let end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}

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

function getAccessToken(req, res, callback) {
  let expiration = new Date( parseFloat( getValueFromCookie('outlook-news-token-expires', req.headers.cookie) ) );

  if (expiration <= new Date()) {
    console.log('TOKEN EXPIRED, REFRESHING');

    let refreshToken = getValueFromCookie('outlook-news-refresh-token', req.headers.cookie);
    authHelper.refreshAccessToken(refreshToken, function(error, newToken) {

      if (error) {
        callback(error, null);
      } else if (newToken) {
        let cookies = [
	  'outlook-news-token=' + newToken.token.access_token + ';Max-Age=4000',
	  'outlook-news-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
	  'outlook-news-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000',
	];

        res.setHeader('Set-Cookie', cookies);
        callback(null, newToken.token.access_token);
      }

    });
  } else {
    callback(null, getValueFromCookie('outlook-news-token', req.headers.cookie) );
  }
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
        let cookies = [
	  'outlook-news-token=' + token.token.access_token + ';Max-Age=4000',
	  'outlook-news-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
	  'outlook-news-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
	  'outlook-news-email=' + email + ';Max-Age=4000'
	];

        res.setHeader('Set-Cookie', cookies);
        res.writeHead(302, {'Location': 'http://localhost:1337/mail'});
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

handle["/mail"] = function(res, req) {
  getAccessToken(req, res, function(error, token) {
    console.log('Token found in cookie: ', token);
    var email = getValueFromCookie('outlook-news-email', req.headers.cookie);
    console.log('Email found in cookie: ', email);
    if (token) {
      
      let query = {
        '$select': 'Subject, ReceivedDateTime, From',
	'$orderby': 'ReceivedDateTime',
	'$top': 15
      };

      outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
      outlook.base.setAnchorMailbox(email);

      outlook.mail.getMessages({token: token, odataParams: query}, function(error, result){
        let page = '';
	
	res.writeHead(200, {'Content-Type': 'text/html'});
        
	if (error) {
          console.log("getMessages returned an error:" + error);

          page = `<p>ERROR: ${error}</p>`;
	} else if (result) {
          console.log("getMessages returned " + result.value.length + ' messages.');
          
	  page = `<div><h1>Your inbox</h1></div>
	  <table>
	    <tr>
	      <th>From</th>
	      <th>Subject</th>
	      <th>Received</th>
	    </tr>`;

	  result.value.forEach(function(message) {
	    console.log('  Subject: ' + message.Subject);

	    let from  = message.From ? message.From.EmailAddress.Name : 'NONE';
	    page += `<tr>
	      <td>${from}</td>
	      <td>${message.Subject}</td>
	      <td>${messge.ReceivedDateTime.toString()}</td>
	    </tr>`;
	  });

	  page += '</table>';
	}
        res.write(page);
        res.end();

      });
    } else {
      res.writeHead(200, {'Content-Type': 'text/html'});
      res.write('<p> No token found in cookie!</p>');
      res.end();
    }
  });
};

server.start(router.route, handle);
