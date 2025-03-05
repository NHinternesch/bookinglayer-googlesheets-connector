// Create the custom menu

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Get Bookinglayer Data')
      .addItem('Get People', 'menuItem1')
      .addItem('Get Bookings', 'menuItem2')
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      fetchAndPastePeopleData();
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      fetchAndPasteBookingData();
}

// Import function 1: Get People

function fetchAndPastePeopleData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = sheet.getSheetByName("Config");
  var peopleSheet = sheet.getSheetByName("People");
  
  if (!configSheet || !peopleSheet) {
    Logger.log("Error: Config or People sheet not found.");
    return;
  }
  
  // Get API key from Config tab, make API request, parse response
  var apiKey = configSheet.getRange("B3").getValue();
  if (!apiKey) {
    Logger.log("Error: API Key not found in Config sheet.");
    return;
  }
  var baseUrl = "https://api.bookinglayer.io/private/persons?page=";
  var limit = 20;
  var page = 1;
  var allData = [];
  
  while (true) {
    var url = baseUrl + page + "&limit=" + limit;
    var options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + apiKey
      }
    };

    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    var persons = json.data;
    
    if (!persons || persons.length === 0) {
      break; // Exit loop if no more persons are returned
    }
    
    allData = allData.concat(persons);
    page++;
  }
  
  if (allData.length === 0) {
    Logger.log("No data fetched.");
    return;
  }
  
  // Prepare data for pasting, clear previous data, paste data to the People tab
  var headers = new Set();
  allData.forEach(person => {
    Object.keys(person).forEach(key => {
      if (typeof person[key] === 'object' && person[key] !== null) {
        Object.keys(person[key]).forEach(nestedKey => {
          headers.add(key + '.' + nestedKey);
        });
      } else {
        headers.add(key);
      }
    });
  });
  
  var headerArray = Array.from(headers);
  
  var values = allData.map(person => {
    return headerArray.map(header => {
      var keys = header.split('.');
      var value = person;
      keys.forEach(k => {
        value = value && value[k] !== undefined ? value[k] : "";
      });
      return value;
    });
  });
  
  peopleSheet.clear();
  
  peopleSheet.getRange(1, 1, 1, headerArray.length).setValues([headerArray]);
  peopleSheet.getRange(2, 1, values.length, headerArray.length).setValues(values);
  
  Logger.log("Data successfully pasted into People sheet.");
}

// Import function 2: Get Bookings

function fetchAndPasteBookingData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = sheet.getSheetByName("Config");
  var bookingsSheet = sheet.getSheetByName("Bookings");

  if (!configSheet || !bookingsSheet) {
    Logger.log("Error: Config or Bookings sheet not found.");
    return;
  }

  // Get API key from Config tab, make API request, parse response
  var apiKey = configSheet.getRange("B3").getValue();
  if (!apiKey) {
    Logger.log("Error: API Key not found in Config sheet.");
    return;
  }

  var baseUrl = "https://api.bookinglayer.io/private/bookings?page=";
  var limit = 20;
  var page = 1;
  var allData = [];

  while (true) {
    var url = baseUrl + page + "&limit=" + limit;
    var options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + apiKey
      }
    };

    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    var bookings = json.data;

    if (!bookings || bookings.length === 0) {
      break; // Exit loop if no more bookings are returned
    }

    allData = allData.concat(bookings);
    page++;
  }

  if (allData.length === 0) {
    Logger.log("No data fetched.");
    return;
  }

  // Prepare data for pasting, clear previous data, paste data to the People tab
  var headers = new Set();
  allData.forEach(booking => {
    Object.keys(booking).forEach(key => {
      if (typeof booking[key] === 'object' && booking[key] !== null && !Array.isArray(booking[key])) {
        Object.keys(booking[key]).forEach(nestedKey => {
          headers.add(key + '.' + nestedKey);
        });
      } else {
        headers.add(key);
      }
    });
  });

  var headerArray = Array.from(headers);

  var values = allData.map(booking => {
    return headerArray.map(header => {
      var keys = header.split('.');
      var value = booking;
      keys.forEach(k => {
        value = value && value[k] !== undefined ? value[k] : "";
      });
      return value;
    });
  });

  bookingsSheet.clear();

  // Paste headers and data
  bookingsSheet.getRange(1, 1, 1, headerArray.length).setValues([headerArray]);
  bookingsSheet.getRange(2, 1, values.length, headerArray.length).setValues(values);

  Logger.log("Data successfully pasted into Bookings sheet.");
}
