function getRentalizerComps(){
  
  try {
    var sheet = getInputSheet();
  } catch(error) {
    Logger.log(error);
    throw "Error in finding the correct sheet for collecting the input data";
  }
  
  var address = sheet.getRange('Address').getValue();
  var bedrooms = sheet.getRange('Bed').getValue();
  var bathrooms = sheet.getRange('Bath').getValue();
  var accommodates = sheet.getRange('Accommodates').getValue();

  
  var addressData = {
	"address": address,
	"bedrooms": bedrooms,
	"bathrooms": bathrooms,
	"accommodates": accommodates,

	"shared_secret": "XXXXXXXX"
  };
  
  var options = {
    'method' : 'post',
    'contentType' : 'application/json',
    'headers': {
      'Accept': 'application/json',
      'Content-Type' : 'application/json'
    }, // API needs these headers, they weren't getting attached because we didn't have headers in the options
    'payload' : JSON.stringify(addressData)
  };
  
  var result = UrlFetchApp.fetch('https://app.rented.com/pricing/airdna/v1/rentalizer/estimate', options);
  result = JSON.parse(result.getContentText());
  
  if (result["error"]) {
    displayDialog();
    throw "Error. See pop-up dialog.";
  }
  
  // Logger.log(result);
  return result;
}



// -----

function getRegionOrCityID(address_lookup) {
  var queryURL = "https://api.airdna.co/client/v1/market/search?access_token=XXXXXXXXXX&term="
  var result = UrlFetchApp.fetch(queryURL + address_lookup);
  result = JSON.parse(result.getContentText());
  
  if (result["items"][result["items"].length-1]["region"] && result["items"][result["items"].length-1]["region"]["id"] > 0) {
    return [result["items"][result["items"].length-1]["region"]["id"],"region"];
  } else if (result["items"][result["items"].length-1]["city"] && result["items"][result["items"].length-1]["city"]["id"] > 0) {
    return [result["items"][result["items"].length-1]["city"]["id"],"city"];
  } else {
    displayDialog();
  }
  
}

 // -----

function getPercentiles(id, bedrooms, address_lookup) {
  if (bedrooms > 4) {
    bedrooms = "5+";
  }

  
  var queryURL = "https://api.airdna.co/client/v1/market/revenue/monthly?" + id[1] + "_id=" + id[0] + "&bedrooms=" + bedrooms + "&room_types=entire_place&start_year=2015&start_month=1&number_of_months=60&access_token=XXXXX";
  
  Logger.log(queryURL);
  var result = UrlFetchApp.fetch(queryURL);
  result = JSON.parse(result.getContentText());
  var json = result;
  writePercentileSheet(json);
 
}

function writePercentileSheet(json) {
//  Logger.log("called writePercentileSheet");
  var sheet = SpreadsheetApp.getActive().getSheetByName('PercentilesPaste');
  
  sheet.getRange('A:Z').clearContent() // Deletes all content except the date formula in the AA column.
  
  var date_range__start_date = json["date_range"]["start_date"];
  var date_range__end_date = json["date_range"]["end_date"];
  var request_info__start_month	= json["request_info"]["start_month"];
  var request_info__start_year = json["request_info"]["start_year"];	
  var request_info__room_types__  = json["request_info"]["room_types"][0];	
  var request_info__currency = json["request_info"]["currency"];	
  var request_info__bedrooms__ = json["request_info"]["bedrooms"][0];	
  var request_info__number_of_months = json["request_info"]["number_of_months"];	
  var request_info__req_type = json["request_info"]["req_type"];	
  var request_info__region_id = json["request_info"]["region_id"];	

  sheet.getRange(1,1).setValue("date_range__start_date");
  sheet.getRange(1,2).setValue("date_range__end_date");
  sheet.getRange(1,3).setValue("request_info__start_month");
  sheet.getRange(1,4).setValue("request_info__start_year");
  sheet.getRange(1,5).setValue("request_info__room_types__");
  sheet.getRange(1,6).setValue("request_info__currency");
  sheet.getRange(1,7).setValue("request_info__bedrooms__");
  sheet.getRange(1,8).setValue("request_info__number_of_months");
  sheet.getRange(1,9).setValue("request_info__req_type");
  sheet.getRange(1,10).setValue("request_info__region_id");
  
  sheet.getRange(2,1).setValue(date_range__start_date);
  sheet.getRange(2,2).setValue(date_range__end_date);
  sheet.getRange(2,3).setValue(request_info__start_month);
  sheet.getRange(2,4).setValue(request_info__start_year);
  sheet.getRange(2,5).setValue(request_info__room_types__);
  sheet.getRange(2,6).setValue(request_info__currency);
  sheet.getRange(2,7).setValue(request_info__bedrooms__);
  sheet.getRange(2,8).setValue(request_info__number_of_months);
  sheet.getRange(2,9).setValue(request_info__req_type);
  sheet.getRange(2,10).setValue(request_info__region_id);
  
  sheet.getRange(1,11).setValue("data__revenue__calendar_months__month");
  sheet.getRange(1,12).setValue("data__revenue__calendar_months__total_revenue");
  sheet.getRange(1,13).setValue("data__revenue__calendar_months__percentiles__25");
  sheet.getRange(1,14).setValue("data__revenue__calendar_months__percentiles__50");
  sheet.getRange(1,15).setValue("data__revenue__calendar_months__percentiles__75");
  sheet.getRange(1,16).setValue("data__revenue__calendar_months__percentiles__90");
  sheet.getRange(1,17).setValue("data__revenue__calendar_months__year");
  sheet.getRange(1,23).setValue("City");
  sheet.getRange(1,24).setValue("Country");
  sheet.getRange(3,23).setValue(json["area_info"]["geom"]["name"]["city"]);
  sheet.getRange(3,24).setValue(json["area_info"]["geom"]["name"]["country"]);
  
  writeCalendarMonths(json, sheet)
  
  
}

function writeCalendarMonths(json, sheet) {
  
  var row = 2
  
  var calendar_months = json["data"]["revenue"]["calendar_months"];
  
  for (var a = 0; a < calendar_months.length; a++) {
    var _25 = calendar_months[a]["percentiles"]["25"];
    var _50 = calendar_months[a]["percentiles"]["50"];
    var _75 = calendar_months[a]["percentiles"]["75"];
    var _90 = calendar_months[a]["percentiles"]["90"];
    var year = calendar_months[a]["year"];
//    Logger.log("calendar_months[year]");
//    Logger.log(calendar_months[3]);

    var month = calendar_months[a]["month"];
    var total_revenue = calendar_months[a]["total_revenue"];
//    Logger.log(calendar_months[a]); 
        
    sheet.getRange(row,11).setValue(month);
    sheet.getRange(row,12).setValue(total_revenue);
    sheet.getRange(row,13).setValue(_25);
    sheet.getRange(row,14).setValue(_50);
    sheet.getRange(row,15).setValue(_75);
    sheet.getRange(row,16).setValue(_90);
    sheet.getRange(row,17).setValue(year);

    row = row + 1; 
  } 
}

 // -----



function fetchAirDNAData() {
  
  Logger.log("Fetch AirDNA Data Called ");


  var currentMonth = getCurrentMonth();
  
  var json = getRentalizerComps();
//  var sheet = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName('RentalizerPaste');
  sheet.clear();
//  Logger.log('sheet');
//  Logger.log(sheet.getSheetName());

  var keys = Object.keys(json).sort();
//  Logger.log('keys');
//  Logger.log(keys);
  
  var last = sheet.getLastColumn();
//  Logger.log('last');
//  Logger.log(last);
  
  var header = sheet.getRange(1, 1).getValues()[0];
  var newCols = [];

 // Write header labels
  sheet.getRange(1,1).setValue("property_details__bathrooms");
  sheet.getRange(1,2).setValue("property_details__bedrooms");                               
  sheet.getRange(1,3).setValue("property_details__zipcode");
  sheet.getRange(1,4).setValue("property_details__accommodates");                               
  sheet.getRange(1,5).setValue("property_details__location__lat");
  sheet.getRange(1,6).setValue("property_details__location__lng");
  sheet.getRange(1,7).setValue("property_details__address");
  sheet.getRange(1,8).setValue("property_details__address_lookup");                             
  sheet.getRange(1,9).setValue("comps__airbnb_property_id"); 
  sheet.getRange(1,10).setValue("comps__stats__|"); 
  
  // sheet.getRange(1,11).setValue("comps__stats__|__" + (getCurrentMonth())); 
    var currentMonth = getCurrentMonth();
  for (var a = 0; a < 12; a++) {

//    Logger.log(currentMonth)
//     if current month = 13, set to 0, else add one
    
    if (currentMonth === 13) {
      sheet.getRange(1,11 + a).setValue("comps__stats__|__" + (1)); 
      currentMonth = 2;
    } else {
      sheet.getRange(1,11 + a).setValue("comps__stats__|__" + (currentMonth)); 
      currentMonth++;
    }
    
  }
  
  sheet.getRange(1,23).setValue("comps__stats__|__ltm");
  sheet.getRange(1,24).setValue("comps__bathrooms");
  sheet.getRange(1,25).setValue("comps__title");
  sheet.getRange(1,26).setValue("comps__cover_img");
  sheet.getRange(1,27).setValue("comps__listing_url");
  sheet.getRange(1,28).setValue("comps__bedrooms");
  sheet.getRange(1,29).setValue("comps__accommodates");
  sheet.getRange(1,30).setValue("comps__location__lat");
  sheet.getRange(1,31).setValue("comps__location_lng");
  sheet.getRange(1,32).setValue("comps__distance_meters");
  sheet.getRange(1,33).setValue("property_stats__|"); 


  
  for (var a = 0; a < 12; a++) {
    
    if (currentMonth === 13) {
      sheet.getRange(1,34 + a).setValue("property_stats__|" + (1)); 
      currentMonth = 2;
    } else {
      sheet.getRange(1,34 + a).setValue("property_stats__|" + (currentMonth)); 
      currentMonth++;
    }
    
  }
  
  sheet.getRange(1,46).setValue("property_stats__|__ltm");
  sheet.getRange(1,47).setValue("permission");

  
  
  sheet.getRange(2,1).setValue(json["property_details"]["bathrooms"]);
  sheet.getRange(2,2).setValue(json["property_details"]["bedrooms"]);
  sheet.getRange(2,3).setValue(json["property_details"]["zipcode"]);
  sheet.getRange(2,4).setValue(json["property_details"]["accommodates"]);
  sheet.getRange(2,5).setValue(json["property_details"]["location"]["lat"]);
  sheet.getRange(2,6).setValue(json["property_details"]["location"]["lng"]);
  sheet.getRange(2,7).setValue(json["property_details"]["address"]);
  sheet.getRange(2,8).setValue(json["property_details"]["address_lookup"]);  

  
  var comps = json["comps"];
  var property_stats = json["property_stats"];
  parseComps(comps, sheet);
  writeRentalizerSummary(property_stats, sheet);
  
  getPercentiles(
    getRegionOrCityID(json["property_details"]["address_lookup"]),
    json["property_details"]["bedrooms"],
    json["property_details"]["address_lookup"]
  );
  
  
//  }
};

 // -----

  function parseComps(comps, sheet) {
    
    // Column 9 is the airbnb property id
    // Column 10 is the name of the stat
    
    var compIndex = 2; // starts on two because we want to start writing in the second row
    
    for (var a = 0; a < Object.keys(comps).length; a++) {

      // Collect all the property details 
      var comps__airbnb_property_id = comps[a]["airbnb_property_id"];
      var comps__bathrooms = comps[a]["bathrooms"];   
      var comps__title = comps[a]["title"];
      var comps__cover_img = comps[a]["cover_img"];
      var comps__listing_url = comps[a]["listing_url"];
      var comps__bedrooms = comps[a]["bedrooms"];
      var comps__accommodates = comps[a]["accommodates"];
      var comps__location__lat = comps[a]["location"]["lat"];
      var comps__location__lng = comps[a]["location"]["lng"];
      var comps__distance_meters = comps[a]["distance_meters"];
      
      // Flatten yearly values for
          // days available
          // adr
          // occupancy
          // revenue_potential
          // revenue
    
      var month_key = parseAnnualData(comps[a]["stats"]["revenue"]);
      var days_available = parseAnnualData(comps[a]["stats"]["days_available"]);
      var adr = parseAnnualData(comps[a]["stats"]["adr"]);
      var occupancy = parseAnnualData(comps[a]["stats"]["occupancy"]);
      var revenue_potential = parseAnnualData(comps[a]["stats"]["revenue_potential"]);
      var revenue = parseAnnualData(comps[a]["stats"]["revenue"]);
      
      // Write out static property details (bedrooms, Airbnb listing, etc.)
      sheet.getRange(compIndex,9).setValue(comps__airbnb_property_id); 
      sheet.getRange(compIndex,24).setValue(comps__bathrooms);
      sheet.getRange(compIndex,25).setValue(comps__title);
      sheet.getRange(compIndex,26).setValue(comps__cover_img);
      sheet.getRange(compIndex,27).setValue(comps__listing_url);
      sheet.getRange(compIndex,28).setValue(comps__bedrooms);
      sheet.getRange(compIndex,29).setValue(comps__accommodates);
      sheet.getRange(compIndex,30).setValue(comps__location__lat);
      sheet.getRange(compIndex,31).setValue(comps__location__lng);
      sheet.getRange(compIndex,32).setValue(comps__distance_meters);

      // Write out all the arrays
      
      for (var x = compIndex; x < compIndex+5; x++) {
        sheet.getRange(compIndex,10).setValue("days_available");
        writePropertyDetails(compIndex, 11, days_available, 1, sheet, 23);

        sheet.getRange(compIndex+1,10).setValue("adr");
        writePropertyDetails(compIndex+1, 11, adr, 1, sheet, 23);
        
        sheet.getRange(compIndex+2,10).setValue("occupancy");
        writePropertyDetails(compIndex+2, 11,  occupancy, 1, sheet, 23);

        sheet.getRange(compIndex+3,10).setValue("revenue_potential");
        writePropertyDetails(compIndex+3, 11, revenue_potential, 1, sheet, 23);

        sheet.getRange(compIndex+4,10).setValue("revenue");
        writePropertyDetails(compIndex+4, 11, revenue, 1, sheet, 23);

      }
     compIndex = compIndex + 5; 
    }
  }

 // -----

function writePropertyDetails(row, column, array, start_month, sheet, ltm_column){
  var monthly = array[0];
  var ltm = array[1];
  var offset = getCurrentMonth() - 1;
  var a = 0;
  
  sheet.getRange(row,ltm_column).setValue(ltm);
  
  for (var x = 0; x < monthly.length; x++) {
    var pointer = (x + offset) % monthly.length;
    sheet.getRange(row,column+a).setValue(monthly[pointer]);
    a = a + 1;
  }

  
}

 // -----

function parseAnnualData(annualData) {
  var dataObj = {};
  var dataList = [];
  var ltm = annualData["ltm"];
  var keys = Object.keys(annualData);
  for (var a = 0; a < keys.length; a++) {
    for (var attrname in annualData[keys[a]]) { 
      dataObj[attrname] = annualData[keys[a]][attrname];
    }
  }
  
//  Logger.log("parseAnnualData annual data");
//  Logger.log(annualData);
//  Logger.log("parseAnnualData combined object");
//  Logger.log(dataObj);
  
  dataList.push(dataObj["1"]);
  dataList.push(dataObj["2"]);
  dataList.push(dataObj["3"]);
  dataList.push(dataObj["4"]);
  dataList.push(dataObj["5"]);
  dataList.push(dataObj["6"]);
  dataList.push(dataObj["7"]);
  dataList.push(dataObj["8"]);
  dataList.push(dataObj["9"]);
  dataList.push(dataObj["10"]);
  dataList.push(dataObj["11"]);
  dataList.push(dataObj["12"]);
//  dataList.push(ltm);

  
  // List returned in order Jan to Dec
  return [dataList, ltm];
}

 // -----

function writeRentalizerSummary(property_stats, sheet) {

  sheet.getRange(2,33).setValue("adr");
  sheet.getRange(3,33).setValue("occupancy");
  sheet.getRange(4,33).setValue("revenue");
  

  var adr = parseAnnualData(property_stats["adr"]);
  writePropertyDetails(2, 34, adr, 1, sheet, 46);

  var occupancy = parseAnnualData(property_stats["occupancy"]);
  writePropertyDetails(3, 34, occupancy, 1, sheet, 46);

  var revenue = parseAnnualData(property_stats["revenue"]);
  writePropertyDetails(4, 34, revenue, 1, sheet, 46);

}
