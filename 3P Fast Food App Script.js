function getPrices() {
  let spreadSheets = SpreadsheetApp.getActiveSpreadsheet();
  let infoSheets = spreadSheets.getSheetByName("CustomerDetails");
  let priceSheets = spreadSheets.getSheetByName("Price");
  let latestRow = infoSheets.getLastRow();
  let latestRow2 = priceSheets.getLastRow();


  let time = infoSheets.getRange(latestRow, 1).getValue();
  let name = infoSheets.getRange(latestRow, 2).getValue();
  let email = infoSheets.getRange(latestRow, 3).getValue();
  let typeOfOrders = infoSheets.getRange(latestRow, 4).getValue();
  let quantities = infoSheets.getRange(latestRow, 5).getValue();
  let pricesRange = infoSheets.getRange(latestRow, 6); // Changed to get the Range object



  let data = priceSheets.getDataRange().getValues();

  // Loop through each row starting from the second row (index 1) if your first row contains headers
  for (let i = 1; i < data.length; i++) {
    let rOrder = data[i][0]; // First column
    let rPrice = data[i][1]; // Second column

    if (typeOfOrders == rOrder) {
      console.log("Same");
      let newPrice = quantities * rPrice;
      // Assuming pricesRange is defined as the range where you want to set the new price
      pricesRange.setValue(newPrice); // Set the value using the Range object
      formatTimestamps()
      getTotalPriceForEachMonth();
      createOrUpdateMonthlyIncomeGraph();
      break; // Exit the loop once the match is found
    } else {
      console.log("Not Same");
    }
  }

}

function formatTimestamps() {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get the range of cells that contain the timestamps (adjust as necessary)
  var range = sheet.getRange("A1:A"); // Adjust the range if needed
  var values = range.getValues();

  // Get the current date
  var currentDate = new Date();

  // Iterate through each cell in the range
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] instanceof Date) {
      var timestamp = values[i][0];
      var cell = range.getCell(i + 1, 1);

      var monthsDifference = getMonthsDifference(currentDate, timestamp);

      if (monthsDifference > 2) {
        cell.setBackground("red");
      } else if (monthsDifference > 1) {
        cell.setBackground("yellow");
      } else if (monthsDifference <= 1 && monthsDifference >= 0) {
        cell.setBackground("green");
      } else {
        cell.setBackground("white");
      }
    }
  }
}

function getMonthsDifference(date1, date2) {
  var year1 = date1.getFullYear();
  var year2 = date2.getFullYear();
  var month1 = date1.getMonth();
  var month2 = date2.getMonth();

  var monthsDifference = (year1 - year2) * 12 + (month1 - month2);

  return monthsDifference;
}

// Set the trigger for the function to run every time the spreadsheet is edited
function createEditTrigger() {
  ScriptApp.newTrigger('formatTimestamps')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}



function getTotalPriceForEachMonth() {
  // Get the active spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Get all data from the sheet
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Object to store month totals
  var monthTotals = {};

  // Iterate through data rows
  data.forEach(function (row) {
    var dateCell = row[0]; // Assuming the date is in the first column
    var priceCell = row[5]; // Assuming the price is in the sixth column

    if (dateCell instanceof Date) {
      var monthKey = dateCell.getFullYear() + '-' + ('0' + (dateCell.getMonth() + 1)).slice(-2);

      if (!monthTotals[monthKey]) {
        monthTotals[monthKey] = 0;
      }

      var price = parseFloat(priceCell);
      if (!isNaN(price)) {
        monthTotals[monthKey] += price;
      }
    }
  });

  // Log the month totals
  Logger.log(monthTotals);

  // Display the month totals in the spreadsheet starting from column H
  var startRow = 1;
  var startColumn = 8; // Column H
  sheet.getRange(startRow, startColumn).setValue("Month");
  sheet.getRange(startRow, startColumn + 1).setValue("Total Price");

  var rowIndex = startRow + 1;
  for (var month in monthTotals) {
    sheet.getRange(rowIndex, startColumn).setValue(month);
    sheet.getRange(rowIndex, startColumn + 1).setValue(monthTotals[month]);
    rowIndex++;
  }
}


function countMonthsInData() {
  // Get the active spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Get all data from the sheet
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Object to store month counts
  var monthCounts = {};

  // Iterate through data rows
  data.forEach(function (row) {
    var dateCell = row[0]; // Assuming the date is in the first column
    if (dateCell instanceof Date) {
      var monthKey = dateCell.getFullYear() + '-' + ('0' + (dateCell.getMonth() + 1)).slice(-2);
      if (!monthCounts[monthKey]) {
        monthCounts[monthKey] = 0;
      }
      monthCounts[monthKey]++;
    }
  });

  // Log the month counts
  Logger.log(monthCounts);
}



function getTotalPriceForEachMonth() {
  // Get the active spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Get all data from the sheet
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Object to store month totals
  var monthTotals = {};

  // Iterate through data rows
  data.forEach(function (row) {
    var dateCell = row[0]; // Assuming the date is in the first column
    var priceCell = row[5]; // Assuming the price is in the sixth column

    if (dateCell instanceof Date) {
      var monthKey = dateCell.getFullYear() + '-' + ('0' + (dateCell.getMonth() + 1)).slice(-2);

      if (!monthTotals[monthKey]) {
        monthTotals[monthKey] = 0;
      }

      var price = parseFloat(priceCell);
      if (!isNaN(price)) {
        monthTotals[monthKey] += price;
      }
    }
  });

  // Log the month totals
  Logger.log(monthTotals);

  // Display the month totals in the spreadsheet starting from column H
  var startRow = 1;
  var startColumn = 9; // Column H
  sheet.getRange(startRow, startColumn).setValue("Month");
  sheet.getRange(startRow, startColumn + 1).setValue("Total Price");

  var rowIndex = startRow + 1;
  for (var month in monthTotals) {
    sheet.getRange(rowIndex, startColumn).setValue(month);
    sheet.getRange(rowIndex, startColumn + 1).setValue(monthTotals[month]);
    rowIndex++;
  }
}

function createOrUpdateMonthlyIncomeGraph() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Find the last row with data in column I
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange("I1:J" + lastRow); // Adjust the range to include all data

  var charts = sheet.getCharts();
  var chartFound = false;

  // Loop through existing charts and update if necessary
  for (var i = 0; i < charts.length; i++) {
    var chart = charts[i];
    var chartPosition = chart.getContainerInfo().getAnchorRow();

    // Check if the chart is positioned just after the last row of data
    if (chartPosition == 1) {
      var newChart = chart.modify()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setOption('title', 'Monthly Income')
        .setOption('hAxis', { title: 'Month' })
        .setOption('vAxis', { title: 'Total Price' })
        .build();

      sheet.updateChart(newChart);
      chartFound = true;
      break;
    }
  }

  // Create a new chart if none was found
  if (!chartFound) {
    var newChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(1, 12, 0, 0) // Position the chart just after the last row of data
      .setOption('title', 'Monthly Income')
      .setOption('hAxis', { title: 'Month' })
      .setOption('vAxis', { title: 'Total Price' })
      .build();

    sheet.insertChart(newChart);
  }
}

function AddItem() {

  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //DEFINE MENU SHEET          
  var invoiceSheet = ss.getSheetByName("INVOICES");
  var itemSheet = ss.getSheetByName("ITEMS");

  //GET NEXT ROW OF INVOICE SHEET
  var lastrowInvoice = invoiceSheet.getLastRow() + 1;

  //GET LAST ROW OF ITEM SHEET
  var lastrowItem = itemSheet.getLastRow();

  // GET VALUE OF PART AND QUANTITY
  var part = invoiceSheet.getRange('B9').getValue();
  var quantity = invoiceSheet.getRange('B10').getValue();

  // GET UNIT PRICE FROM ITEM SHEET
  for (var i = 2; i <= lastrowItem; i++) {
    if (part == itemSheet.getRange(i, 1).getValue()) {
      var unitCost = itemSheet.getRange(i, 2).getValue();
    }
  }

  // POPULATE INVOICE SHEET
  invoiceSheet.getRange(lastrowInvoice, 1).setValue(part);
  invoiceSheet.getRange(lastrowInvoice, 2).setValue(quantity);
  invoiceSheet.getRange(lastrowInvoice, 3).setValue(unitCost).setNumberFormat('"RM"#,###.00');

}


function createInvoice() {
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //DEFINE MENU SHEET          
  var invoiceSheet = ss.getSheetByName("INVOICES");
  var customerSheet = ss.getSheetByName("CUSTOMERS");
  var settingSheet = ss.getSheetByName("SETTINGS");
  var printSheet = ss.getSheetByName("PRINT");

  //GET VALUES
  var name = invoiceSheet.getRange(2, 2).getValue();
  var payable_to = invoiceSheet.getRange(3, 2).getValue();
  var project_name = invoiceSheet.getRange(4, 2).getValue();
  var invoice_number = settingSheet.getRange(1, 2).getValue();
  var next_invoice_number = invoice_number + 1;
  settingSheet.getRange(1, 2).setValue(next_invoice_number);

  var due_date = invoiceSheet.getRange(5, 2).getValue();
  var note = invoiceSheet.getRange(6, 2).getValue();
  var adjustment = invoiceSheet.getRange(7, 2).getValue();


  // GET CUSTOER LAST ROW
  var lastrowCustomer = customerSheet.getLastRow();

  // GET CUSTOMER FIELDS
  for (var i = 2; i <= lastrowCustomer; i++) {
    if (name == customerSheet.getRange(i, 1).getValue()) {
      var companyName = customerSheet.getRange(i, 2).getValue();
      var streetAddress = customerSheet.getRange(i, 3).getValue();
      var city = customerSheet.getRange(i, 4).getValue();
      var state = customerSheet.getRange(i, 5).getValue();
      var zip = customerSheet.getRange(i, 6).getValue();
    }
  }

  // SET INVOICE DATE
  var currentDate = new Date();
  var currentMonth = currentDate.getMonth() + 1;
  var currentYear = currentDate.getFullYear();
  var date = currentMonth.toString() + '/' + currentDate.getDate().toString() + '/' + currentYear.toString();


  // GET LAST ROW OF PRINT SHEET
  var lastrowPrint = printSheet.getLastRow();

  // FIND HOW MANY ITEMS ROWS TO DELETE
  var x_count = 0
  for (var v = 19; v <= lastrowPrint; v++) {

    if (printSheet.getRange(v, 2).getValue() != 'Notes:') {
      x_count++;
    }
    else {
      break;
    }
  }

  //Logger.log(x_count);

  var lastrowPrint = 19 + x_count;

  //Logger.log(lastrowPrint);

  // DELETE ITEMS ROWS FROM INVOICE
  if ((lastrowPrint - 19) != 0) {
    printSheet.deleteRows(19, lastrowPrint - 19);
  }

  // SET VALUES ON INVOICE
  printSheet.getRange('B9').setValue('Submitted to ' + date).setFontFamily('Roboto').setFontSize(12).setFontWeight("bold").setFontColor("#e01b84");
  printSheet.getRange('B12').setValue(name).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B13').setValue(companyName).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B14').setValue(streetAddress).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B15').setValue(city + ', ' + state + ' ' + zip).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('D12').setValue(payable_to).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('D15').setValue(project_name).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('F12').setValue(invoice_number).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('F15').setValue(due_date).setFontFamily('Roboto').setFontSize(10).setFontColor("3e01b84");

  printSheet.getRange('C19').setValue(note).setFontFamily('Roboto').setFontSize(10).setFontColor("#E01B84");
  printSheet.getRange('G20').setValue(adjustment).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");


  // GET LAST ROW OF INVOICE SHEET
  var lastrowInvoice = invoiceSheet.getLastRow();

  var z = 0;
  var subTotal = 0;
  for (var y = 15; y <= lastrowInvoice; y++) {
    //INSERT ROW ON PRINT SHEET
    printSheet.insertRowsAfter(18 + z, 1);

    //GET ITEM VALUES FROM INVOICE SHEET
    var part = invoiceSheet.getRange(y, 1).getValue();
    var quantity = invoiceSheet.getRange(y, 2).getValue();
    var unitPrice = invoiceSheet.getRange(y, 3).getValue();

    // PRICE TOTALS
    var totalPrice = quantity * unitPrice;
    subTotal = subTotal + totalPrice;

    // POPULATE TOTALS ON PRINT SHEET
    printSheet.getRange(18 + z + 1, 2).setValue(part).setFontFamily('Roboto').setFontSize(10).setFontColor("black");
    printSheet.getRange(18 + z + 1, 5).setValue(quantity).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
    printSheet.getRange(18 + z + 1, 6).setValue(unitPrice).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
    printSheet.getRange(18 + z + 1, 7).setValue(totalPrice).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");

    z++;
  }

  // SET TOTAL
  printSheet.getRange(18 + z + 1, 7).setValue(subTotal).setNumberFormat("$#,###.00").setFontFamily('Roboto').setFontSize(10).setFontColor("black");

  var totalInvoice = subTotal + adjustment;

  // CALL INVOICE LOG
  InvoiceLog(invoice_number, name, date, due_date, totalInvoice)

}

function InvoiceLog(invoice_number, name, date, due_date, totalInvoice) {

  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //DEFINE INVOICE LOG SHEET          
  var invoiceLogSheet = ss.getSheetByName("INVOICE LOG");

  //GET LAST ROW OF INVOICE LOG SHEET
  var nextRowInvoice = invoiceLogSheet.getLastRow() + 1;

  //POPULATE INVOICE LOG
  invoiceLogSheet.getRange(nextRowInvoice, 1).setValue(invoice_number);
  invoiceLogSheet.getRange(nextRowInvoice, 2).setValue(name);
  invoiceLogSheet.getRange(nextRowInvoice, 3).setValue(date);
  invoiceLogSheet.getRange(nextRowInvoice, 4).setValue(due_date);
  invoiceLogSheet.getRange(nextRowInvoice, 5).setValue(totalInvoice);



}

function ClearInvoice() {
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //DEFINE INVOICE SHEET          
  var invoiceSheet = ss.getSheetByName("INVOICES");


  //SET VALUES TO NOTHING
  invoiceSheet.getRange(2, 2).setValue("");
  invoiceSheet.getRange(3, 2).setValue("");
  invoiceSheet.getRange(4, 2).setValue("");
  invoiceSheet.getRange(5, 2).setValue("");
  invoiceSheet.getRange(6, 2).setValue("");
  invoiceSheet.getRange(7, 2).setValue("");
  invoiceSheet.getRange(8, 2).setValue("");
  invoiceSheet.getRange(9, 2).setValue("");
  invoiceSheet.getRange(10, 2).setValue("");

  //CLEAR ITEMS
  invoiceSheet.getRange("A15:C1000").clear();

}


