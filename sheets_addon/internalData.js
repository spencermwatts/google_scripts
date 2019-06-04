
var apiKey = "XXX";

// InternalData.gs takes data from our internal services and writes them to the google sheets doc.

// This function is where the magic happens. 
function writeInternalCompsToSheet(){
  try {
  // get comp listings from JSON response
  var comparable_listings = getInternalCompData();
  // get all the header data
  var headers = collectHeaders(comparable_listings).sort()
  
  // find the right sheet to write data & clear.
  var sheet = SpreadsheetApp.getActive().getSheetByName('Internal Comps');
  sheet.clearContents();
  
  // write the comp details headers (property id, beds, baths, etc.)
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // write the details about each property 
  writeCompValues(headers, comparable_listings, sheet);
  
  // write the tables of historical performance for each property
  writeHistoricalPerformance(comparable_listings);
  } catch(error) {
  }
  
}


function getInternalCompData(){
//  try {
    var sheet = getInputSheet();
//  } catch(error) {
//    throw "This Google Doc needs a sheet called 'Inputs' which contains the named ranges for address/bedroom/bathrooms/accomodates";
//  }
  var address = sheet.getRange('Address').getValue();
  var bedrooms = sheet.getRange('Bed').getValue();
  var bathrooms = sheet.getRange('Bath').getValue();
  var accommodates = sheet.getRange('Accommodates').getValue();
  
  var addressData = {
	"address": address,
	"bedrooms": bedrooms,
	"bathrooms": bathrooms,
	"accommodates": accommodates,
  };
  
  
  var options = {
    'method' : 'post',
    'contentType' : 'application/json',
    'headers': {
      'Accept': 'application/json',
      'Content-Type' : 'application/json'
    }, 
    'payload' : JSON.stringify(addressData)
  };
  
  var result = UrlFetchApp.fetch("http://XXX.us-east-2.compute.amazonaws.com:8080/market_info?key=" + apiKey, options);
  result = JSON.parse(result.getContentText());
  
  sheet.getRange(1,1).setValue(result.market_info_response_id);
//  Logger.log(result.market_info_response_id);
  
  if (result["error"]) {
    displayDialog();
    throw "Error. See pop-up dialog.";
  }
  
  return result["comparable_listings"];
}

// This function takes arrays of objects and find distinct headers.
// For example, we need to find the distinct metrics a time series of objects has.
function collectHeaders(objs){ 
  var headers = [];
  for (var i = 0; i < objs.length; ++i){ // for each object 
    for (var key in objs[i]) { // for each key in that object
      if (headers.indexOf(key) == -1){
        headers.push(key);
      }
    }
  } 
  return headers;
}

// This function takes a list of headers (i.e., data points) and 
// comps and writes out the data for each comp.
function writeCompValues(headers, comparable_listings, sheet) {
  for (var i = 0; i < comparable_listings.length; ++i){
    for (var e = 0; e < headers.length; ++e){
      // i + 2 as (1) the row range must be 1 or more and (2) we already have a header row.
      // e + 1 as (1) the column range must be 1 or more.
      sheet.getRange(i+2,e+1).setValue(comparable_listings[i][headers[e]]); 
    }
  }
} 

// This function takes the historical data and writes it to a table.
function writeHistoricalPerformance(comparable_listings) {
  
  // need to clear this new sheet vefore use
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('Internal Comps Performance');
  sheet.clearContents();
  
  
  var row = 1;
  var column = 3; 
  // starting at third column as we want the 1st to contain 
  // the property ID and the 2nd to contain the metric name
  
  // Get the column (months) and row (metrics) labels for the performance table
  var data = getHistoricalTableData(comparable_listings);
  var monthly_objects = data[0];
  var metrics = collectHeaders(monthly_objects);
  var months = data[1];
  

  // Write just the months
  sheet.getRange(row, column, 1, months.length).setValues([months]);
  row++;

  
  // Now write the data for those months & metrics.
  for (var i = 0; i < comparable_listings.length; ++i) {    // For each comp
    if (comparable_listings[i].historical_performance.length == 0) { // Skip properties with no data.
      continue;
    }
    for (var hd = 0; hd < comparable_listings[i].historical_performance.length; ++hd) { // For each month of a given comp
      var column_offset = months.indexOf(comparable_listings[i].historical_performance[hd].reporting_month)
      for (var m = 0; m < metrics.length; ++m) {
        sheet.getRange(row + m + 1, 1).setValue(comparable_listings[i].property_id); // Write property ID
        sheet.getRange(row + m + 1, 2).setValue(metrics[m]); // Write Metric
        sheet.getRange(row + m + 1, column + column_offset).setValue(comparable_listings[i].historical_performance[hd][metrics[m]]); // Write data
        }
    }
    row = row + metrics.length;
  }
}

// This function gets the row and column labels for our historic data table. 
// It gets the monthly_objects (i.e., metrics) as well as the months of data we have.
function getHistoricalTableData(historic_data) {
  var monthly_objects = [];
  var months = [];
  var months_a = [];
  for (var i = 0; i < historic_data.length; ++i) {
    
    for (var m = 0; m < historic_data[i].historical_performance.length; ++m) {
      monthly_objects.push(historic_data[i].historical_performance[m]);
      if (months.indexOf(historic_data[i].historical_performance[m]["reporting_month"]) == -1){
        months.push(historic_data[i].historical_performance[m]["reporting_month"]);
      }  
    }
  }
      return [monthly_objects, months.sort()];
}



// ------------------------------------------------------------------------------------

// This function sends the data back to John's data science service!
function emitJSON() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Data');
  var output = {};
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  
  var parents_key = sheet.getSheetValues(1, 2, numRows, 1)
  var children_key = sheet.getSheetValues(1, 3, numRows, 1)
  var values = sheet.getSheetValues(1, 5, numRows, 1)
  
  // Create a giant JSON blob! 
  for (var i = 0; i < numRows; ++i) {
    output[sheet.getRange(i+1,1).getValue()] = sheet.getRange(i+1,4).getValue(); 
    // i+1 because you can't have a 0 row
    // Getting value from column 4 bc that should be the unformatted #
  }
  
  var options = {
    'method' : 'post',
    'contentType' : 'application/json',
    'headers': {
      'Accept': 'application/json',
      'Content-Type' : 'application/json'
    }, 
    'payload' : JSON.stringify(output)
  };
  
  var result = UrlFetchApp.fetch("http://xxx.us-east-2.compute.amazonaws.com:8080/appraise?key=" + apiKey, options);
  
  //  Logger.log(Object.keys(output));
  //  Logger.log(JSON.stringify(output));
  //  Logger.log(JSON.stringify(output));
  
  //  for (var i = 0; i < numRows; ++i) {
  //    if (parents_key[i].join().length) { // If there is a value at all
  //      Logger.log(parents_key[i]);
  //    }
  //  }
}
