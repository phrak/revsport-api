// Global variables
var PORTAL_NAME = "1st-balcombe-heigh";
var SEASON_ID = "35822";
var USERNAME = "tuidam1";
var PASSWORD = "XXXPASSWORDXXX";
var DEBUG_OUTPUT_FILENAME = "Get_ScoutHub_Members_Debug.html";
var CSV_OUTPUT_FILENAME = "Get_ScoutHub_Members_Data.csv";

function importScoutHubMembers() {
  console.log("Starting importScoutHubMembers function");

  if (PASSWORD === "XXXPASSWORDXXX") {
    throw new Error("Please change the default password in the script before running it.");
  }

  console.log("PORTAL_NAME:", PORTAL_NAME);
  console.log("SEASON_ID:", SEASON_ID);
  console.log("USERNAME:", USERNAME);
  console.log("Login URL:", "https://client.revolutionise.com.au/?clientName=" + PORTAL_NAME + "&page=/" + PORTAL_NAME + "/");

  var api = new RevSportAPI(PORTAL_NAME);

  if (api.loginOld(USERNAME, PASSWORD)) {
    console.log("Login successful, proceeding to fetch members");
    var csvData = api.fetchMembers();

    if (csvData) {
      try {
        writeDebugFilesToFolder(api.debugHtml, csvData);
        var data = Utilities.parseCsv(csvData);
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        sheet.clear();
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        console.log("Data import completed successfully");
      } catch (e) {
        console.error("Error parsing or writing CSV data:", e.message);
        console.log("CSV data preview:", csvData.substring(0, 500));
      }
    } else {
      console.error("Failed to fetch members data");
    }
  } else {
    console.error("Login failed, unable to proceed");
  }
}

function writeDebugFilesToFolder(debugHtml, csvData) {
  const folderName = "DEBUG";
  let folderIterator = DriveApp.getFoldersByName(folderName);
  let folder;

  if (folderIterator.hasNext()) {
    folder = folderIterator.next();
    console.log(`Folder '${folderName}' already exists.`);
  } else {
    folder = DriveApp.createFolder(folderName);
    console.log(`Folder '${folderName}' created.`);
  }

  const debugFile = folder.createFile(DEBUG_OUTPUT_FILENAME, debugHtml, MimeType.HTML);
  console.log(`Debug HTML file written to folder: ${debugFile.getUrl()}`);

  const csvFile = folder.createFile(CSV_OUTPUT_FILENAME, csvData, MimeType.CSV);
  console.log(`CSV file written to folder: ${csvFile.getUrl()}`);
}

if (typeof RevSportAPI === 'undefined') {
  class RevSportAPI {
    constructor(portal) {
      this.portalName = portal;
      this.session = {};
      this.debugHtml = "";
    }

    loginOld(username, password) {
      console.log("Starting loginOld method");
      var url = "https://client.revolutionise.com.au/?clientName=" + this.portalName + "&page=/" + this.portalName + "/";

      try {
        var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
        console.log("Login page fetched. Status code:", response.getResponseCode());

        if (response.getResponseCode() !== 200) {
          console.error("Failed to fetch login page. Status code:", response.getResponseCode());
          return false;
        }

        var html = response.getContentText();
        console.log("Login page HTML length:", html.length);

        var initialCookies = response.getAllHeaders()['Set-Cookie'];
        console.log("Initial cookies:", initialCookies);

        var tokenMatch = html.match(/<input[^>]*name="_csrf"[^>]*value="([^"]*)"[^>]*>/);
        if (!tokenMatch) {
          console.error("CSRF token not found in the HTML");
          console.log("HTML preview:", html.substring(0, 1000));
          return false;
        }
        var token = tokenMatch[1];
        console.log("CSRF token found:", token);

        var loginUrl = "https://client.revolutionise.com.au/" + this.portalName + "/scripts/login/client/";
        console.log("Submitting login form to:", loginUrl);
        var options = {
          'method': 'post',
          'payload': {
            '_csrf': token,
            'redirect': '/' + this.portalName + '/',
            'user': username,
            'password': password
          },
          'followRedirects': false,
          'muteHttpExceptions': true,
          'headers': {
            'Cookie': initialCookies
          }
        };

        var loginResponse = UrlFetchApp.fetch(loginUrl, options);
        console.log("Login response status:", loginResponse.getResponseCode());

        var headers = loginResponse.getAllHeaders();
        console.log("Response headers:", JSON.stringify(headers));

        if (headers['Set-Cookie']) {
          this.session.cookies = Array.isArray(headers['Set-Cookie'])
            ? headers['Set-Cookie'].join('; ')
            : headers['Set-Cookie'];
          console.log("Cookies set successfully:", this.session.cookies);
        } else {
          console.error("No cookies received in the response");
        }

        var redirectUrl = loginResponse.getHeaders()['Location'];
        while (redirectUrl) {
          console.log("Following redirect to:", redirectUrl);
          var redirectOptions = {
            'method': 'get',
            'followRedirects': false,
            'muteHttpExceptions': true,
            'headers': {
              'Cookie': this.session.cookies
            }
          };
          var redirectResponse = UrlFetchApp.fetch(redirectUrl, redirectOptions);
          console.log("Redirect response status:", redirectResponse.getResponseCode());

          if (redirectResponse.getAllHeaders()['Set-Cookie']) {
            this.session.cookies = Array.isArray(redirectResponse.getAllHeaders()['Set-Cookie'])
              ? redirectResponse.getAllHeaders()['Set-Cookie'].join('; ')
              : redirectResponse.getAllHeaders()['Set-Cookie'];
            console.log("Updated cookies after redirect:", this.session.cookies);
          }

          redirectUrl = redirectResponse.getHeaders()['Location'];
        }

        this.writeDebugLog("login_success.html", "Login successful. Final response headers: " + JSON.stringify(redirectResponse.getAllHeaders()));
        return true;
      } catch (e) {
        console.error("Error in loginOld:", e.message, e.stack);
        return false;
      }
    }

    fetchMembers() {
      console.log("Starting fetchMembers method");
      var url = "https://client.revolutionise.com.au/" + this.portalName + "/members/reports/";

      var cookieHeader = this.session.cookies;

      var options = {
        'method': 'get',
        'headers': {
          'Cookie': cookieHeader
        },
        'muteHttpExceptions': true
      };

      console.log("Sending request with cookies:", cookieHeader);

      try {
        var response = UrlFetchApp.fetch(url, options);
        console.log("Reports page fetched. Status code:", response.getResponseCode());
        console.log("Response headers:", JSON.stringify(response.getAllHeaders()));

        var html = response.getContentText();
        this.debugHtml = html;

        console.log("HTML content preview:", html.substring(0, 1000));

        if (html.indexOf('<title>Generate members report') !== -1) {
          console.log('Success: The page title indicates a successful fetch.');
        } else if (html.indexOf('<title>Please Log In') !== -1) {
          console.error('Error: The page title indicates a login failure.');
          this.writeDebugLog("login_failure.html", html);

          // Re-login mechanism
          console.log("Attempting to re-login...");
          if (this.loginOld(USERNAME, PASSWORD)) {
            console.log("Re-login successful, retrying to fetch members");
            return this.fetchMembers(); // Recursive call to retry fetching members
          } else {
            console.error("Re-login failed, unable to proceed");
            return null;
          }
        }

        this.writeDebugLog("reports_page.html", html);

      } catch (e) {
        console.error("Error in fetchMembers:", e.message, e.stack);
        return null;
      }
    }

    writeDebugLog(filename, content) {
      const folderName = "DEBUG";
      let folder;
      let folderIterator = DriveApp.getFoldersByName(folderName);

      if (folderIterator.hasNext()) {
        folder = folderIterator.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }

      const file = folder.createFile(filename, content, MimeType.HTML);
      console.log(`Debug log written to: ${file.getUrl()}`);
    }
  }
}
