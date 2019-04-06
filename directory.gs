function listAllUsers() {
  var pageToken, page;
  do {
    page = AdminDirectory.Users.list({
      domain: 'ett-twello.nl',
      orderBy: 'givenName',
      maxResults: 500,
      pageToken: pageToken
    });
    var users = page.users;
    if (users) {
      for (var i = 0; i < users.length; i++) {
        var user = users[i];
        Logger.log('%s (%s) %s %s', user.name.fullName, user.primaryEmail, user.emails, user.lastLoginTime);
      }
    } else {
      Logger.log('No users found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}