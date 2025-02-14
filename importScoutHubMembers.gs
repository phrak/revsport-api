// Global variables
var PORTAL_NAME = "1st-your-group-url";
var SEASON_ID = "35822";
var USERNAME = "XXXUSERNAMEXXX";
var PASSWORD = "XXXPASSWORDXXX";
var DEBUG_OUTPUT_FILENAME = "Get_ScoutHub_Members_Debug.html";
var CSV_OUTPUT_FILENAME = "Get_ScoutHub_Members_Data.csv";

function importScoutHubMembers() {
  console.log("Starting importScoutHubMembers function");
  
  // Check if the password is the default one
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
        // Write debug HTML and CSV to DEBUG folder
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

  // Check if the folder exists
  let folderIterator = DriveApp.getFoldersByName(folderName);
  let folder;

  if (folderIterator.hasNext()) {
    folder = folderIterator.next();
    console.log(`Folder '${folderName}' already exists.`);
  } else {
    // Create the folder if it doesn't exist
    folder = DriveApp.createFolder(folderName);
    console.log(`Folder '${folderName}' created.`);
  }

  // Write the debug HTML file
  const debugFile = folder.createFile(DEBUG_OUTPUT_FILENAME, debugHtml, MimeType.HTML);
  console.log(`Debug HTML file written to folder: ${debugFile.getUrl()}`);

  // Write the CSV file
  const csvFile = folder.createFile(CSV_OUTPUT_FILENAME, csvData, MimeType.CSV);
  console.log(`CSV file written to folder: ${csvFile.getUrl()}`);
}

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
      // Step 1: Fetch the login page
      console.log("Fetching login page:", url);
      var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
      console.log("Login page fetched. Status code:", response.getResponseCode());
      
      if (response.getResponseCode() !== 200) {
        console.error("Failed to fetch login page. Status code:", response.getResponseCode());
        return false;
      }
      
      var html = response.getContentText();
      console.log("Login page HTML length:", html.length);
      
      // Log all cookies received from the initial page load
      var initialCookies = response.getAllHeaders()['Set-Cookie'];
      console.log("Initial cookies:", initialCookies);
      
      // Step 2: Extract CSRF token
      var tokenMatch = html.match(/<input[^>]*name="_csrf"[^>]*value="([^"]*)"[^>]*>/);
      if (!tokenMatch) {
        console.error("CSRF token not found in the HTML");
        console.log("HTML preview:", html.substring(0, 1000));
        return false;
      }
      var token = tokenMatch[1];
      console.log("CSRF token found:", token);
      
      // Step 3: Submit the login form
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
          'Cookie': initialCookies // Include initial cookies in login request
        }
      };
      
      var loginResponse = UrlFetchApp.fetch(loginUrl, options);
      console.log("Login response status:", loginResponse.getResponseCode());
      
      var headers = loginResponse.getAllHeaders();
      console.log("Response headers:", JSON.stringify(headers));
      
      if (headers['Set-Cookie']) {
        this.session.cookies = headers['Set-Cookie'];
        console.log("Cookies set successfully:", this.session.cookies);
      } else {
        console.error("No cookies received in the response");
      }
      
      // Follow all redirects
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
          this.session.cookies = redirectResponse.getAllHeaders()['Set-Cookie'];
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
    var options = {
      'method': 'get',
      'headers': {
        'Cookie': this.session.cookies
      },
      'muteHttpExceptions': true
    };
    
    console.log("Sending request with cookies:", this.session.cookies);
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      console.log("Reports page fetched. Status code:", response.getResponseCode());
      console.log("Response headers:", JSON.stringify(response.getAllHeaders()));
      
      var html = response.getContentText();
      this.debugHtml = html;  // Store the HTML for debug output
      
      // New debug check for page title
      if (html.indexOf('<title>Generate members report') !== -1) {
        console.log('Success: The page title indicates a successful fetch.');
      } else if (html.indexOf('<title>Please Log In') !== -1) {
        console.error('Error: The page title indicates a login failure.');
        this.writeDebugLog("login_failure.html", html);
        return null;
      }
      
      // Write debug log for reports page
      this.writeDebugLog("reports_page.html", html);
      
      var tokenMatch = html.match(/<input[^>]*name="_csrf"[^>]*value="([^"]*)"[^>]*>/);
      if (!tokenMatch) {
        console.error("CSRF token not found in the HTML");
        return null;
      }
      var token = tokenMatch[1];
      console.log("CSRF token found");
      
      // Extract all data field display names and form field names
      var formMatch = html.match(/<form[^>]*id="membersDownload"[^>]*>([\s\S]*?)<\/form>/);
      if (!formMatch) {
        console.error("Form with ID 'membersDownload' not found in the HTML content.");
        return null;
      }
      var formHtml = formMatch[1];

      var fields = {};
      var checkboxMatches = formHtml.match(/<input[^>]*type="checkbox"[^>]*name="([^"]*)"[^>]*/g);
      if (checkboxMatches) {
        for (var i = 0; i < checkboxMatches.length; i++) {
          var name = checkboxMatches[i].match(/name="([^"]*)"/)[1];
          fields[name] = '1';
          if (name === 'last_updated') break;
        }
      }
      console.log("Fields extracted:", Object.keys(fields).length);
      
      var data = {
        '_csrf': token,
        'file_format': 'csv',
        'season_id': SEASON_ID,
        'filterby': '',
        'orderby': 'nationalMemberID',
        'direction': 'asc',
        'nameorder': 'split',
        'addressformat': 'together'
      };
      
      for (var field in fields) {
        data[field] = '1';
      }
      
      // Include hidden fields
      var hiddenMatches = formHtml.match(/<input[^>]*type="hidden"[^>]*>/g);
      if (hiddenMatches) {
        for (var i = 0; i < hiddenMatches.length; i++) {
          var name = hiddenMatches[i].match(/name="([^"]*)"/)[1];
          var value = hiddenMatches[i].match(/value="([^"]*)"/)[1];
          data[name] = value;
        }
      }
      
      // Debugging: Print the data being sent in the POST request
      console.log("POST data being sent:");
      for (var key in data) {
        console.log(key + ": " + data[key]);
      }
      
      var downloadUrl = "https://client.revolutionise.com.au/" + this.portalName + "/reports/members/download/";
      var downloadOptions = {
        'method': 'post',
        'payload': data,
        'headers': {
          'Cookie': this.session.cookies
        },
        'muteHttpExceptions': true
      };
      
      var csvResponse = UrlFetchApp.fetch(downloadUrl, downloadOptions);
      console.log("CSV data fetched. Status code:", csvResponse.getResponseCode());
      console.log("Content type:", csvResponse.getHeaders()['Content-Type']);
      
      var content = csvResponse.getContentText();
      console.log("Content length:", content.length);
      
      // Log the first 500 characters of the content
      console.log("Content preview:", content.substring(0, 500));
      
      // Check if the content starts with expected CSV headers
      if (!content.trim().startsWith("National Member ID,")) {
        console.error("Unexpected content format. First 100 characters:", content.substring(0, 100));
        return null;
      }
      
      return content;
    } catch (e) {
      console.error("Error in fetchMembers:", e.message);
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
