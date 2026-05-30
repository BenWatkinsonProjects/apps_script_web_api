function getOrders() {
  const CK = '???';
  const CS = '???';
  
  const baseUrl = "???";
  const dateFilter = "?after=" + "2026-04-24T00:00:00";
  const pagination = "&per_page=100";
  const statusFilter = "&status=processing";
  const url = baseUrl + dateFilter + pagination + statusFilter;

  const options = {
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(CK + ":" + CS)
    }
  };

  try {
    let response = UrlFetchApp.fetch(url, options);
    let orders = JSON.parse(response.getContentText());
    
    let ordersData = [[
      "Order ID",
      "Customer ID",
      "Order Status",
      "First Name",
      "Last Name",
      "Where did you hear about us?",
      "Invoice Email",
      "Company",
      "Address 1",
      "Address 2",
      "City",
      "County",
      "Postcode",
      "VAT",
      "Total (inc. VAT)",
      "Date Ordered"
    ]];

    orders.forEach(function(order) {
      let sourceOfHearing = getMeta("_wc_other/picts/how-did-you-hear-about-us"); 

      ordersData.push([
        order.id,
        order.customer_id,
        order.status,
        order.billing.first_name,
        order.billing.last_name,
        sourceOfHearing,
        order.billing.email,
        order.billing.company,
        order.billing.address_1,
        order.billing.address_2,    
        order.billing.city,
        order.billing.state,
        order.billing.postcode,
        order.total_tax,
        order.total,
        order.date_created
      ]);
    });

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("???");
    sheet.clearContents();

    sheet.getRange(10, 1, ordersData.length, ordersData[0].length).setValues(ordersData);

    hideEmptyRows(sheet);

    SpreadsheetApp.flush();
    const ui = SpreadsheetApp.getUi();
    ui.alert('Data refreshed - Please wait a few moments for orders to appear... ', ui.ButtonSet.OK);

  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}  

function hideEmptyRows(sheet) {
  const data = sheet.getDataRange().getValues();
  
  sheet.showRows(1, sheet.getMaxRows());

  let startRow = -1;
  let rowCount = 0;

  for (let i = 1; i < data.length; i++) {
    const isRowEmpty = (data[i][1] === "" || data[i][1] === null);

    if (isRowEmpty) {
      if (startRow === -1) {
        startRow = i + 1;
      }
      rowCount++;
    } else {
      if (startRow !== -1) {
        sheet.hideRows(startRow, rowCount);
        startRow = -1;
        rowCount = 0;
      }
    }
  }

  if (startRow !== -1) {
    sheet.hideRows(startRow, rowCount);
  }
}
