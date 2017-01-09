/*
 * Copyright (c) Microsoft All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var request = require('request');
var Q = require('q');
var https = require('https');


// The graph module object.
var graph = {};

// @name getUsers
// @desc Makes a request to the Microsoft Graph for all users in the tenant.
graph.getUsers = function (token) {
  var deferred = Q.defer();

  // Make a request to get all users in the tenant. Use $select to only get
  // necessary values to make the app more performant.
  request.get('https://graph.microsoft.com/v1.0/users?$select=id,displayName', {
    auth: {
      bearer: token
    }
  }, function (err, response, body) {
    var parsedBody = JSON.parse(body);

    if (err) {
      deferred.reject(err);
    } else if (parsedBody.error) {
      deferred.reject(parsedBody.error.message);
    } else {
      // The value of the body will be an array of all users.
      deferred.resolve(parsedBody.value);
    }
  });

  return deferred.promise;
};

// @name createEvent
// @desc Creates an event on each user's calendar.
// @param token The app's access token.
// @param users An array of users in the tenant.
graph.createEvent = function (token, users, r) {
  var i;
  var startTime;
  var endTime;
  var newEvent;
  var id = '86ce0077-bf25-4063-a722-e9d143490b34';
    // The new event will be 30 minutes and take place tomorrow at the current time.
    startTime = new Date();
    startTime.setDate(startTime.getDate() + 1);
    endTime = new Date(startTime.getTime() + 30 * 60000);

    // These are the fields of the new calendar event.
    newEvent = {
      Subject: 'Microsoft Graph API discussion',
      Location: {
        DisplayName: "Joe's office"
      },
      Start: {
        DateTime: startTime,
        TimeZone: 'PST'
      },
      End: {
        DateTime: endTime,
        TimeZone: 'PST'
      },
      Body: {
        Content: 'Let\'s discuss this awesome API.',
        ContentType: 'Text'
      }
    };

    // Add an event to the current user's calendar.
    request.post({
      url: 'https://graph.microsoft.com/v1.0/users/' + id + '/events',
      headers: {
        'content-type': 'application/json',
        authorization: 'Bearer ' + token,
        displayName: 'Eric Halsey'
      },
      body: JSON.stringify(newEvent)
    }, function (err, response, body) {
      var parsedBody;
      var displayName;
      if (err) {
        console.error('>>> Application error: ' + err);
      } else {
        parsedBody = JSON.parse(body);
        displayName = response.request.headers.displayName;

        if (parsedBody.error) {
          var errorMsg;
          if (parsedBody.error.code === 'RequestBroker-ParseUri') {
            errorMsg =
              '>>> Error creating an event for ' + displayName +
              '. Most likely due to this user having a MSA instead of an Office 365 account.'
            console.error(errorMsg);
            r.tellWithCard("Error:" + errorMsg, "Hello World", "Hello World!");
          } else {
            errorMsg = 
              '>>> Error creating an event for ' + displayName + '.' + parsedBody.error.message
            console.error(errorMsg);
            r.tellWithCard("Error:" + errorMsg, "Hello World", "Hello World!");
          }
        } else {
          console.log('>>> Successfully created an event on ' + displayName + "'s calendar.");
          r.tellWithCard("Calendar item created!", "Hello World", "Hello World!");
        }
      }
    });
};

// @name createItem
// @desc Creates an item in a SP List.
// @param token The app's access token.
graph.createItem = function (token, callback) {
  mailBody = '{}';
  var outHeaders = {
    'Content-Type': 'application/json',
    Authorization: 'Bearer ' + token,
    'Content-Length': mailBody.length
  };
  var options = {
    host: 'graph.microsoft.com',
    path: '/beta/sharepoint/sites/6c1838eb-37d3-4c25-9f63-b097e52bc7dc,d4fdfd5b-9add-4a9c-bbbc-daeafbc62833/lists/3e5523f7-3059-4759-b0e4-12aa6d3b041e/items/',
    method: 'POST',
    headers: outHeaders
  };

  // Set up the request
  var post = https.request(options, function (response) {
    var body = '';
    response.on('data', function (d) {
      body += d;
    });
    response.on('end', function () {
      var error;
      if (response.statusCode === 201) {
        callback(null);
      } else {
        error = new Error();
        error.code = response.statusCode;
        error.message = response.statusMessage;
        // The error body sometimes includes an empty space
        // before the first character, remove it or it causes an error.
        body = body.trim();
        error.innerError = JSON.parse(body).error;
        // Note: If you receive a 500 - Internal Server Error
        // while using a Microsoft account (outlok.com, hotmail.com or live.com),
        // it's possible that your account has not been migrated to support this flow.
        // Check the inner error object for code 'ErrorInternalServerTransientError'.
        // You can try using a newly created Microsoft account or contact support.
        callback(error);
      }
    });
  });

  // write the outbound data to it
  post.write(mailBody);
  // we're done!
  post.end();

  post.on('error', function (e) {
    callback(e);
  });
};

module.exports = graph;