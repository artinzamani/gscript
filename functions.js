//sets the value of a given range in a given spreadsheet
function set_value(spreadsheet, sheet_name, value, range) {
    let sheet = SpreadsheetApp.openByUrl(spreadsheet).getSheetByName(sheet_name);
    sheet.getRange(range[0], range[1]).setValue(value);
}

//finds the first empty row based on the requested header
function first_blank_row(spreadsheet, sheet_name, header) {
    var sheet = SpreadsheetApp.openByUrl(spreadsheet).getSheetByName(sheet_name);
    var col = get_col_index(spreadsheet, sheet_name, header);
    var row = 1;
    
    while (sheet.getRange(row, col).getValue() != "")
    {
        row++;
    }
    
    return row;
}

//finds the occurrence of value in sheet based on the header
function find_range(spreadsheet, sheet_name, header, value) {
    var sheet = SpreadsheetApp.openByUrl(spreadsheet).getSheetByName(sheet_name);
    var column = get_col_index(spreadsheet, sheet_name, header);
    var row = 2;
    
    while (sheet.getRange(row, column).getValue() != value)
    {
        row++;
    }
    
    return [row, column];
}

//looks up the value, then returns the range of the cell with the desired header corresponding to that row
function lookup_range(spreadsheet, sheet_name, lookup_header, value, requested_header) {
    var lookup_cell = find_range(spreadsheet, sheet_name, lookup_header, value);
    var lookup_col = lookup_cell[1];
    var lookup_row = lookup_cell[0];
    var requested_col = get_col_index(spreadsheet, sheet_name, requested_header);
    
    return [lookup_row, requested_col];
}

//looks up the value, then inserts the value in the corresponding column for that row
function set(spreadsheet, sheet_name, lookup_header, value, insertion_header, insertion_value) {
    var sheet = SpreadsheetApp.openByUrl(spreadsheet).getSheetByName(sheet_name);
    var range = lookup_range(spreadsheet, sheet_name, lookup_header, value, insertion_header);
    sheet.getRange(range[0], range[1]).setValue(insertion_value);
}

//finds the column index of the requested header
function get_col_index(spreadsheet, sheet_name, header) {
    var sheet = SpreadsheetApp.openByUrl(spreadsheet).getSheetByName(sheet_name);
    var column = 1;
    var row = 1;
    while (sheet.getRange(row, column).getValue() != header)
    {
        column++;
    }
    
    return column;
}

//logs gregorian date
function register_date() {
    var today = new Date();
    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
    return date;
}

//logs jalali date
function register_date_persian() {
    var today = new Date();
    var persianDate = gregorian_to_jalali(today.getFullYear(), parseInt(today.getMonth()) + 1, parseInt(today.getDate()));
    var hours = today.getHours();
    var minutes = today.getMinutes();
    var seconds = today.getSeconds();
    var time = hours + ":" + minutes + ":" + seconds;
    return persianDate[0].toString() + '-' + persianDate[1].toString() + '-' + persianDate[2].toString() + ' ' + time;
}

//convert date to array
function date_to_array(date) {
    date = String(date);
    return date.split('/');
}

//converts gregorian date to jalali
function gregorian_to_jalali(gy, gm, gd) {
    var g_d_m, jy,jm,jd, gy2,days;
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
    if (gy > 1600) {
        jy = 979;
        gy -= 1600;
    } else {
        jy = 0;
        gy -= 621;
    }
    gy2 = (gm > 2) ? (gy + 1) : gy;
    days = (365 * gy) + (parseInt((gy2 + 3) / 4)) - (parseInt((gy2 + 99) / 100)) +(parseInt((gy2 + 399) / 400)) - 80 + gd + g_d_m[gm - 1];
    jy += 33 * (parseInt(days / 12053)); 
    days %= 12053;
    jy += 4 * (parseInt(days / 1461));
    days %= 1461;
    if (days > 365) {
        jy +=parseInt((days - 1) / 365);
        days = (days - 1) % 365;
    }
    jm = (days < 186) ? 1 + parseInt(days / 31) : 7 + parseInt((days - 186) / 30);
    jd = 1 + ((days < 186) ? (days % 31) : ((days - 186) % 30));
    return [jy, jm, jd];
}

//converts jalali date to gregorian
function jalali_to_gregorian(jy, jm, jd) {
    var sal_a, gy, gm, gd, days, v;
    if (jy > 979) {
        gy = 1600;
        jy -= 979;
    } else {
        gy = 621;
    }
    days = (365 * jy) + ((parseInt(jy / 33)) * 8) +(parseInt(((jy % 33) + 3) / 4)) + 78 + jd + ((jm < 7) ? (jm - 1) * 31 : ((jm - 7) * 30) + 186);
    gy += 400 * (parseInt(days / 146097));
    days %= 146097;
    if(days > 36524) {
        gy += 100 * (parseInt(--days / 36524));
        days %= 36524;
        if(days >= 365) days++;
    }
    gy += 4 * (parseInt(days / 1461));
    days %= 1461;
    if(days > 365) {
        gy += parseInt((days - 1) / 365);
        days = (days - 1) % 365;
    }
    gd = days + 1;
    sal_a = [0, 31, ((gy % 4 === 0 && gy % 100 !== 0) || (gy % 400 === 0)) ? 29 : 28 ,31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    for (gm = 0; gm < 13; gm++) {
        v = sal_a[gm];
        if (gd <= v) break;
        gd -= v;
    }
    return [gy,gm,gd]; 
}

function send_sms(to, text) {
    var url = "https://rest.payamak-panel.com/api/SendSMS/SendSMS";
    var payload = {
      "username": "username",
      "password": "password",
      "to": to,
      "from": "from",
      "text": text
    };
    var options = {
      "method": "POST",
      "payload": payload,
      "muteHttpExceptions": true
    };
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() == 200) {
      
      var params = JSON.parse(response.getContentText());
      Logger.log(params.name);
      Logger.log(params.blog);
    }
}
  
function send_email(email, subject, text) {
    GmailApp.sendEmail(email, subject, text);
}

function send_request(method = "get", url, payload) {
    // sends an HTTP request and returns the response.
    // the method in string: get, post put, ...
    // the url
    // payload as json
    var options = {
    "method": method.toUpperCase(),
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() == 200)
    {
      return JSON.parse(response.getContentText());
    }
}

// adds menu to the ui.
function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .createMenu('Menu Name')
        .addItem('Menu Item', 'callback function')
        .addToUi();
}