/**
 * Connect and fetch Salesforce data via OAuth
 */
function queryDataFromSalesforce() {
  // Read OAuth consumer key / secret of this client app from script properties, 
  // which can be issued from Salesforce's remote access setting in advance.
  var sfConsumerKey = ScriptProperties.getProperty("sfConsumerKey");
  var sfConsumerSecret = ScriptProperties.getProperty("sfConsumerSecret");
  if (!sfConsumerKey || !sfConsumerSecret) {
    Browser.msgBox("Register Salesforce OAuth Consumer Key and Secret in Script Properties");
    return;
  }

  // Register new OAuth service, named "salesforce"
  // For OAuth endpoint information, see help doc in Salesforce.
  // https://na7.salesforce.com/help/doc/en/remoteaccess_oauth_1_flows.htm
  var oauth = UrlFetchApp.addOAuthService("salesforce");
  oauth.setAccessTokenUrl("https://login.salesforce.com/_nc_external/system/security/oauth/AccessTokenHandler");
  oauth.setRequestTokenUrl("https://login.salesforce.com/_nc_external/system/security/oauth/RequestTokenHandler");
  oauth.setAuthorizationUrl("https://login.salesforce.com/setup/secur/RemoteAccessAuthorizationPage.apexp?oauth_consumer_key="+encodeURIComponent(sfConsumerKey));
  oauth.setConsumerKey(sfConsumerKey);
  oauth.setConsumerSecret(sfConsumerSecret);

  // Convert OAuth1 access token to Salesforce sessionId (mostly equivalent to OAuth2 access token)
  var sessionLoginUrl = "https://login.salesforce.com/services/OAuth/u/21.0";
  var options = { method : "POST", oAuthServiceName : "salesforce", oAuthUseToken : "always" };
  var result = UrlFetchApp.fetch(sessionLoginUrl, options);
  var txt = result.getContentText();
  var accessToken = txt.match(/<sessionId>([^<]+)/)[1];
  var serverUrl = txt.match(/<serverUrl>([^<]+)/)[1];
  var instanceUrl = serverUrl.match(/^https?:\/\/[^\/]+/)[0];
  
  // Query account data from Salesforce, using REST API with OAuth2 access token.
  var fields = "Id,Name,Type,BillingState,BillingCity,BillingStreet";
  var soql = "SELECT "+fields+" FROM Account LIMIT 100";
  var queryUrl = instanceUrl + "/services/data/v21.0/query?q="+encodeURIComponent(soql);
  var response = UrlFetchApp.fetch(queryUrl, { method : "GET", headers : { "Authorization" : "OAuth "+accessToken } });
  var queryResult = Utilities.jsonParse(response.getContentText());

  // Render query result to Spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  sheet.setFrozenRows(1);

  // Render all field names in header row.
  var cell = sheet.getRange('a1');
  fields = fields.split(',');
  fields.forEach(function(field, j){ cell.offset(0, j).setValue(field) })

  // Render result records into cells
  queryResult.records.forEach(function(record, i) {
    fields.forEach(function(field, j) { cell.offset(i+1, j).setValue(record[field]) });
  });

}
