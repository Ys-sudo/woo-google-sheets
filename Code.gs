// Updated code to fetch WooCommerce orders with pagination and additional fields
function start_syncv2() {
    var sheet_name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    fetch_orders(sheet_name);
}

function fetch_orders(sheet_name) {
    var ck = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B4").getValue();
    var cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B5").getValue();
    var website = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B3").getValue();
    var manualDate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B6").getValue(); // Set your order start date in spreadsheet in cell B6
    var m = new Date(manualDate).toISOString();
    var headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(ck + ':' + cs)
    };

    var perPage = 100;
    var page = 1;
    var allOrders = [];

    while (true) {
        var surl = website + "/wp-json/wc/v3/orders?consumer_key=" + ck + "&consumer_secret=" + cs + "&after=" + m + "&per_page=" + perPage + "&page=" + page;
        var options = {
            "method": "GET",
            "headers": headers,
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "muteHttpExceptions": true,
        };

        var result = UrlFetchApp.fetch(surl, options);
        var headers2 = result.getAllHeaders();
        Logger.log(headers2);

        if (result.getResponseCode() == 200) {
            var params = JSON.parse(result.getContentText());
            if (params.length === 0) {
                break; // No more orders to fetch
            }
            allOrders = allOrders.concat(params);
            page++; // Move to the next page for next iteration
        } else {
            Logger.log('Error fetching URL: ' + result);
            break; // Exit the loop on error
        }
    }

    // Process allOrders array containing all fetched orders
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var temp = doc.getSheetByName(sheet_name);

    var arrayLength = allOrders.length;
    Logger.log(arrayLength);

    for (var i = 0; i < arrayLength; i++) {
        var container = [];

        // Billing Details
        container.push(allOrders[i]["billing"]["first_name"]);
        container.push(allOrders[i]["billing"]["last_name"]);
        container.push(allOrders[i]["billing"]["address_1"] + " " + allOrders[i]["billing"]["postcode"] + " " + allOrders[i]["billing"]["city"]);
        container.push(allOrders[i]["billing"]["phone"]);
        container.push(allOrders[i]["billing"]["email"]);

        // Shipping Details
        container.push(allOrders[i]["shipping"]["first_name"] + " " + allOrders[i]["shipping"]["last_name"] + " " + allOrders[i]["shipping"]["address_1"] + " " + allOrders[i]["shipping"]["postcode"] + " " + allOrders[i]["shipping"]["city"] + " " + allOrders[i]["shipping"]["country"]);

        // Order Details
        container.push(allOrders[i]["customer_note"]);
        container.push(allOrders[i]["payment_method_title"]);
        container.push(allOrders[i]["total"]);
        container.push(allOrders[i]["total_tax"]);
        container.push(allOrders[i]["discount_total"]);
        container.push(allOrders[i]["refunded_total"]);

        // Line Items
        var lineItems = "";
        var totalQuantity = 0;
        for (var j = 0; j < allOrders[i]["line_items"].length; j++) {
            var item = allOrders[i]["line_items"][j]["name"];
            var quantity = allOrders[i]["line_items"][j]["quantity"];
            lineItems += quantity + " x " + item + ",\n";
            totalQuantity += quantity;
        }
        container.push(lineItems);
        container.push(totalQuantity);

        // Refunds
        var refunds = "";
        for (var k = 0; k < allOrders[i]["refunds"].length; k++) {
            var refundReason = allOrders[i]["refunds"][k]["reason"];
            var refundAmount = allOrders[i]["refunds"][k]["total"];
            refunds += refundAmount + " - " + refundReason + ",\n";
        }
        container.push(refunds);

        // Order Meta
        container.push(allOrders[i]["number"]);
        container.push(allOrders[i]["date_created"]);
        container.push(allOrders[i]["date_modified"]);
        container.push(allOrders[i]["status"]);
        container.push(allOrders[i]["order_key"]);
        container.push(allOrders[i]["currency"]);
        container.push(allOrders[i]["cart_tax"]);
        container.push(allOrders[i]["shipping_total"]);
        container.push(allOrders[i]["shipping_tax"]);
        container.push(allOrders[i]["customer_user_agent"]);
        container.push(allOrders[i]["customer_ip_address"]);

        // Append the row to the sheet
        temp.appendRow(container);
    }

    removeDuplicates(sheet_name);
}

function removeDuplicates(sheet_name) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(sheet_name);
    var data = sheet.getDataRange().getValues();
    var newData = [];

    for (var i in data) {
        var duplicate = false;
        for (var j in newData) {
            if (data[i].join() == newData[j].join()) {
                duplicate = true;
                break;
            }
        }
        if (!duplicate) {
            newData.push(data[i]);
        }
    }

    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}
