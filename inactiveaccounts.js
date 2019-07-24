/* 
-- No Impressions Accounts Checker --
Purpose: This script is run at the Accounts Summary (CID) level to identify accounts that have no impressions on the day of running.
Scope: Account Summary
*/


/* 
Setup Part 1: Specify Account ID's
If you wish to specify only certain Account ID's, then please specify by setting the below equal too an array of account ID's.
As an example, replace the below with: let accountIds = ['XID1','XID2','XID3'] 
*/

let accountIds;

/* 
Setup Part 2: Email Alert Preference
Set emailAlert to true if you wish this script to email an alert, set this to false if you do not. 
If you set this to false, please ignore Setup Parts 3 and 4. 
Note: with no email alert you will need to check the logs once the script runs for the results.
*/

const emailAlert = true;

/* 
Setup Part 3: Google Services API Credentials
Please follow Option 2 in https://docs.microsoft.com/en-us/advertising/scripts/examples/authenticating-with-google-services to receive your clientId, 
clientSecret, and refreshToken. It is recommended you create a separate Gmail account specifically for running Microsoft Advertising Scripts.
*/ 

const credentials = {
  clientId: '',
  clientSecret: '',
  refreshToken: ''
};

/* 
Setup Part 4: Email Recipients and Subject
Please update emailRecipients to include those you wish to be included in the email alert and emailSubject to the email subject of your choice.
*/ 

const emailRecipients = ['email1@example.com','email2@example.com'];
const emailSubject = 'Microsoft Advertising | Inactive Accounts Alert';

/* 
 DO NOT EDIT THE BELOW CODE THIS POINT
*/

function main() {
  let accounts = getAccounts();
  if (accounts) {
    let inactiveAccounts = ``;
    while (accounts.hasNext()){
      let currentAccount = accounts.next();
      let accountImpressions = currentAccount.getStats().getImpressions();
      inactiveAccounts += `<tr><td align="left">${currentAccount.getName()}</td><td align="left">${currentAccount.getAccountNumber()}</td><td align="left">${accountImpressions}</td></tr>`;
      Logger.log(`Potential Inactive Account: ${currentAccount.getName()} || ${currentAccount.getAccountNumber()} || Impressions: ${accountImpressions}`);
    }

    if (emailAlert) {
        const emailMessage = `<html><body><h1>Microsoft Advertising - Inactive Accounts Alert</h1><p>The below accounts have had no impressions today and are potentially inactive:</p> <table border = "1" width="95%" style="border-collapse:collapse;"><tr><th align="left">Account Name</th><th align="left">Account Id</th><th align="left">Impressions Today</th></tr>${inactiveAccounts}</table></body></html>`;
        sendEmail(emailRecipients, emailSubject, emailMessage);
    } 

  } else {
    Logger.log('All accounts have impressions so far today.')
  }
}

function sendEmail(emailRecipients, emailSubject, emailMessage){
  let gmailApi = GoogleApis.createGmailService(credentials);
  let email = [`To: ${emailRecipients.join(',')}`,`Subject: ${emailSubject}`,'Content-Type: text/html','',`${emailMessage}`].join('\n');
  try {
      let sendResponse = gmailApi.users.messages.send({ userId: 'me' }, { raw: Base64.encode(email) });
  }
  catch(e){
      Logger.log(`There was an issue trying to send the email: ${e}`);
  }
}

function getAccounts(){
  let accounts;
  if (accountIds != undefined){
    accounts = AccountsApp.accounts().withAccountNumbers(accountIds).withCondition("Impressions = 0").forDateRange("TODAY").get();
  } else {
    accounts = AccountsApp.accounts().withCondition("Impressions = 0").forDateRange("TODAY").get();
  }
  
  if(accounts.hasNext()){
    return accounts;
  } else {
    return false;
  } 
}

let GoogleApis;
(function (GoogleApis) { 
  function createGmailService(credentials) {
    return createService("https://www.googleapis.com/discovery/v1/apis/gmail/v1/rest", credentials);
  }
  GoogleApis.createGmailService = createGmailService;
 
  // Creation logic based on https://developers.google.com/discovery/v1/using#usage-simple
  function createService(url, credentials) {
    var content = UrlFetchApp.fetch(url).getContentText();
    var discovery = JSON.parse(content);
    var baseUrl = discovery['rootUrl'] + discovery['servicePath'];
    var accessToken = getAccessToken(credentials);
    var service = build(discovery, {}, baseUrl, accessToken);
    return service;
  }
 
  function createNewMethod(method, baseUrl, accessToken) {
    return (urlParams, body) => {
      var urlPath = method.path;
      var queryArguments = [];
      for (var name in urlParams) {
        var paramConfg = method.parameters[name];
        if (!paramConfg) {
          throw `Unexpected url parameter ${name}`;
        }
        switch (paramConfg.location) {
          case 'path':
            urlPath = urlPath.replace('{' + name + '}', urlParams[name]);
            break;
          case 'query':
            queryArguments.push(`${name}=${urlParams[name]}`);
            break;
          default:
            throw `Unknown location ${paramConfg.location} for url parameter ${name}`;
        }
      }
      var url = baseUrl + urlPath;
      if (queryArguments.length > 0) {
        url += '?' + queryArguments.join('&');
      }
      var httpResponse = UrlFetchApp.fetch(url, { contentType: 'application/json', method: method.httpMethod, payload: JSON.stringify(body), headers: { Authorization: `Bearer ${accessToken}` }, muteHttpExceptions: true });
      var responseContent = httpResponse.getContentText();
      var responseCode = httpResponse.getResponseCode();
      var parsedResult;
      try {
        parsedResult = JSON.parse(responseContent);
      } catch (e) {
        parsedResult = false;
      }
      var response = new Response(parsedResult, responseContent, responseCode);
      if (responseCode >= 200 && responseCode <= 299) {
        return response;
      }
      throw response;
    }
  }
 
  function Response(result, body, status) {
    this.result = result;
    this.body = body;
    this.status = status;
  }
  Response.prototype.toString = function () {
    return this.body;
  }
 
  function build(discovery, collection, baseUrl, accessToken) {
    for (var name in discovery.resources) {
      var resource = discovery.resources[name];
      collection[name] = build(resource, {}, baseUrl, accessToken);
    }
    for (var name in discovery.methods) {
      var method = discovery.methods[name];
      collection[name] = createNewMethod(method, baseUrl, accessToken);
    }
    return collection;
  }
 
  function getAccessToken(credentials) {
    if (credentials.accessToken) {
      return credentials.accessToken;
    }
    var tokenResponse = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v4/token', { method: 'post', contentType: 'application/x-www-form-urlencoded', muteHttpExceptions: true, payload: { client_id: credentials.clientId, client_secret: credentials.clientSecret, refresh_token: credentials.refreshToken, grant_type: 'refresh_token' } });    
    var responseCode = tokenResponse.getResponseCode(); 
    var responseText = tokenResponse.getContentText(); 
    if (responseCode >= 200 && responseCode <= 299) {
      var accessToken = JSON.parse(responseText)['access_token'];
      return accessToken;
    }    
    throw responseText;  
  }
})(GoogleApis || (GoogleApis = {}));
// Base64 implementation from https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/master/lib/msal-core/src/Utils.ts
class Base64 {
  static encode(input) {
    const keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
    let output = "";
    let chr1, chr2, chr3, enc1, enc2, enc3, enc4;
    var i = 0;
    input = this.utf8Encode(input);
    while (i < input.length) {
      chr1 = input.charCodeAt(i++);
      chr2 = input.charCodeAt(i++);
      chr3 = input.charCodeAt(i++);
      enc1 = chr1 >> 2;
      enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
      enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
      enc4 = chr3 & 63;
      if (isNaN(chr2)) {
        enc3 = enc4 = 64;
      }
      else if (isNaN(chr3)) {
        enc4 = 64;
      }
      output = output + keyStr.charAt(enc1) + keyStr.charAt(enc2) + keyStr.charAt(enc3) + keyStr.charAt(enc4);
    }
    return output.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }
  static utf8Encode(input) {
    input = input.replace(/\r\n/g, "\n");
    var utftext = "";
    for (var n = 0; n < input.length; n++) {
      var c = input.charCodeAt(n);
      if (c < 128) {
        utftext += String.fromCharCode(c);
      }
      else if ((c > 127) && (c < 2048)) {
        utftext += String.fromCharCode((c >> 6) | 192);
        utftext += String.fromCharCode((c & 63) | 128);
      }
      else {
        utftext += String.fromCharCode((c >> 12) | 224);
        utftext += String.fromCharCode(((c >> 6) & 63) | 128);
        utftext += String.fromCharCode((c & 63) | 128);
      }
    }
    return utftext;
  }
}
