var auth = require('./auth');
var graph = require('./graph');

exports.handler = (event, context, callback) => {
  // Get an access token for the app.
  auth.getAccessToken().then(function (token) {
    // Get all of the users in the tenant.
    graph.getUsers(token)
      .then(function (users) {
        // Create an event on each user's calendar.
        graph.createEvent(token, users);
      }, function (error) {
        console.error('>>> Error getting users: ' + error);
      });
  }, function (error) {
    console.error('>>> Error getting access token: ' + error);
  });
  callback(null, 'Hello from Lambda');
};

