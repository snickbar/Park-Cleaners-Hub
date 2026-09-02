/**
 * Main entry point: Serves the homepage (Index.html)
 */
function doGet() {
  const output = HtmlService.createHtmlOutputFromFile('Index');
  output.setTitle('Park Cleaners Invoice Manager');
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

/**
 * Page Navigation Helper
 */
function loadPage(pageName) {
  return HtmlService.createHtmlOutputFromFile(pageName).getContent();
}

/**
 * Fetch customer list from the 'Customers' sheet.
 */
function getCustomerList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).map(row => row[0]).filter(String);
}

/**
 * Fetch service list and prices from the 'Services' sheet.
 */
function getServiceList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Services");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).map(row => {
    return {
      service: row[0],
      price: Number(row[1]) || 0
    };
  }).filter(item => item.service);
}

/**
 * Finalizes the invoice, logs to History, builds a real PDF matching exact format,
 * and saves it into the "Invoices" Drive folder.
 */
function finalizeInvoice(invoiceData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName("History");
  
  if (!historySheet) {
    historySheet = ss.insertSheet("History");
    historySheet.appendRow(["Invoice #", "Date", "Customer", "Items", "Total", "PDF Link"]);
  }
  
  // Calculate Invoice #
  const lastRow = historySheet.getLastRow();
  const invoiceNum = lastRow > 1 ? 1000 + (lastRow - 1) : 1080;
  
  // Format dates (MM/DD/YYYY)
  const rawDate = invoiceData.date ? new Date(invoiceData.date + "T00:00:00") : new Date();
  
  const serviceDateFormatted = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
  const dateIssuedFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");

  // 1. Locate or create the "Invoices" folder in Google Drive
  const folderName = "Invoices";
  const folders = DriveApp.getFoldersByName(folderName);
  let invoiceFolder;
  if (folders.hasNext()) {
    invoiceFolder = folders.next();
  } else {
    invoiceFolder = DriveApp.createFolder(folderName);
  }

  // 2. Build HTML table rows matching exact layout
  let tableRowsHtml = "";
  invoiceData.items.forEach(function(item) {
    tableRowsHtml += `
      <tr>
        <td style="border: 2px solid #000000; padding: 12px 10px; font-weight: bold;">${serviceDateFormatted}</td>
        <td style="border: 2px solid #000000; padding: 12px 10px; font-weight: bold;">${item.service}</td>
        <td style="border: 2px solid #000000; padding: 12px 10px; font-weight: bold;">${item.notes || ''}</td>
        <td style="border: 2px solid #000000; padding: 12px 10px; font-weight: bold; text-align: left;">${item.quantity}</td>
        <td style="border: 2px solid #000000; padding: 12px 10px; font-weight: bold;">$${item.price.toFixed(2)}</td>
        <td style="border: 2px solid #000000; padding: 12px 10px; font-weight: bold;">$${item.total.toFixed(2)}</td>
      </tr>
    `;
  });

  // 3. Build full printable HTML layout
  const fullPdfHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @page {
          size: letter portrait;
          margin: 0.5in;
        }
        body {
          font-family: Arial, Helvetica, sans-serif;
          margin: 30px;
          color: #000000;
          background: #ffffff;
        }
        
        /* HEADER SECTION */
        .company-header {
          text-align: center;
          margin-bottom: 40px;
        }
        .company-name {
          font-size: 38px;
          font-weight: 800;
          letter-spacing: -0.5px;
          margin-bottom: 6px;
        }
        .company-sub {
          font-size: 16px;
          font-weight: bold;
          line-height: 1.4;
        }

        /* METADATA SECTION */
        .meta-container {
          width: 100%;
          margin-bottom: 20px;
        }
        .bill-to-label {
          font-size: 18px;
          font-weight: bold;
        }
        .bill-to-name {
          font-size: 20px;
          font-weight: bold;
          margin-top: 4px;
        }
        .meta-right {
          text-align: right;
          font-size: 20px;
          font-weight: bold;
        }

        .divider {
          border: none;
          border-top: 2px solid #777777;
          margin-top: 15px;
          margin-bottom: 25px;
        }

        /* INVOICE TABLE */
        .invoice-table {
          width: 100%;
          border-collapse: collapse;
          border: 2px solid #000000;
        }
        .invoice-table th {
          background-color: #f2f2f2;
          border: 2px solid #000000;
          padding: 12px 10px;
          text-align: left;
          font-size: 16px;
          font-weight: 800;
          text-transform: uppercase;
        }
        .invoice-table td {
          font-size: 15px;
        }

        /* TOTAL SECTION */
        .total-container {
          text-align: right;
          margin-top: 35px;
          font-size: 28px;
          font-weight: 800;
          letter-spacing: 0.5px;
        }

        /* FOOTER SECTION */
        .footer-container {
          text-align: center;
          margin-top: 60px;
          font-size: 20px;
          font-weight: bold;
          line-height: 1.4;
        }
      </style>
    </head>
    <body>

      <!-- HEADER -->
      <div class="company-header">
        <div class="company-name">Park Cleaners</div>
        <div class="company-sub">
          Professional Cleaning & Tailoring Services<br>
          97 Pehle Ave, Saddle Brook, NJ 07663<br>
          201-388-1044 &bull; park.cleaners@yahoo.com
        </div>
      </div>

      <!-- BILL TO & INVOICE DETAILS -->
      <table class="meta-container">
        <tr>
          <td style="vertical-align: top;">
            <div class="bill-to-label">Bill To</div>
            <div class="bill-to-name">${invoiceData.customer}</div>
          </td>
          <td style="vertical-align: top;" class="meta-right">
            <div>Invoice #: ${invoiceNum}</div>
            <div style="margin-top: 4px;">Date Issued: ${dateIssuedFormatted}</div>
          </td>
        </tr>
      </table>

      <hr class="divider">

      <!-- ITEMS TABLE -->
      <table class="invoice-table">
        <thead>
          <tr>
            <th style="width: 17%;">SERVICE DATE</th>
            <th style="width: 31%;">DESCRIPTION</th>
            <th style="width: 22%;">NOTES</th>
            <th style="width: 10%;">QTY</th>
            <th style="width: 10%;">PRICE</th>
            <th style="width: 10%;">TOTAL</th>
          </tr>
        </thead>
        <tbody>
          ${tableRowsHtml}
        </tbody>
      </table>

      <!-- GRAND TOTAL -->
      <div class="total-container">
        TOTAL: $${invoiceData.total.toFixed(2)}
      </div>

      <!-- FOOTER -->
      <div class="footer-container">
        Thank you for your business!<br>
        Barry Barone
      </div>

    </body>
    </html>
  `;

  // 4. Create real PDF file inside Drive folder
  const blob = Utilities.newBlob(fullPdfHtml, 'text/html', `Invoice_${invoiceNum}_${invoiceData.customer}.html`);
  const pdfFile = invoiceFolder.createFile(blob.getAs('application/pdf'));
  
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = pdfFile.getUrl();

  // 5. Append record to History sheet
  const itemsSummary = invoiceData.items.map(item => {
    const noteText = item.notes ? ` (${item.notes})` : '';
    return `${item.quantity}x ${item.service}${noteText} [$${item.price.toFixed(2)}]`;
  }).join("; ");

  historySheet.appendRow([
    invoiceNum,
    dateIssuedFormatted,
    invoiceData.customer,
    itemsSummary,
    invoiceData.total,
    pdfUrl
  ]);

  return {
    success: true,
    invoiceNum: invoiceNum,
    pdfUrl: pdfUrl
  };
}

/**
 * Fetches all invoice history records and unique customer list for searching/sorting.
 */
function getInvoiceHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("History");
  const customerSheet = ss.getSheetByName("Customers");
  
  if (!historySheet) return { invoices: [], customers: [] };
  
  // Use getDisplayValues() instead of getValues() to read exact formatted cell text
  const data = historySheet.getDataRange().getDisplayValues();
  let invoices = [];
  
  if (data.length > 1) {
    invoices = data.slice(1).map(row => {
      return {
        invoiceNum: row[0] ? row[0].toString() : '',
        date: row[1] ? row[1].toString() : '',
        customer: row[2] ? row[2].toString() : '',
        items: row[3] ? row[3].toString() : '',
        total: row[4] ? row[4].toString() : '',
        pdfUrl: row[5] ? row[5].toString() : ''
      };
    });
  }

  let customers = [];
  if (customerSheet) {
    const custData = customerSheet.getDataRange().getValues();
    if (custData.length > 1) {
      customers = custData.slice(1).map(r => r[0]).filter(String);
    }
  }

  return {
    invoices: invoices,
    customers: customers
  };
}

/**
 * Fetch all services from the 'Services' sheet for the Services Manager page.
 */
function getServicesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Services");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const services = [];
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Ensure Service name isn't empty
      services.push({
        rowIndex: i + 1, // Sheet row reference
        service: data[i][0].toString().trim(),
        price: data[i][1] ? data[i][1].toString().trim() : "0.00"
      });
    }
  }
  return services;
}

/**
 * Save a new service or update an existing one in the 'Services' sheet.
 */
function saveService(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Services");

  if (!sheet) {
    sheet = ss.insertSheet("Services");
    sheet.appendRow(["Service", "Price"]);
  }

  const serviceName = data.service.trim();
  const priceVal = parseFloat(data.price) || 0;
  const formattedPrice = priceVal.toFixed(2);

  if (data.rowIndex) {
    // Update existing row
    sheet.getRange(data.rowIndex, 1, 1, 2).setValues([[serviceName, formattedPrice]]);
  } else {
    // Add new row
    sheet.appendRow([serviceName, formattedPrice]);
  }

  return { success: true };
}

/**
 * Delete a service row from the 'Services' sheet.
 */
function deleteService(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Services");
  if (sheet && rowIndex) {
    sheet.deleteRow(rowIndex);
  }
  return { success: true };
}

/**
 * Fetch all customers from the 'Customers' sheet for the Customers Manager page.
 */
function getCustomersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const customers = [];
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Ensure Customer Name isn't empty
      customers.push({
        rowIndex: i + 1, // Sheet row reference
        name: data[i][0].toString().trim()
      });
    }
  }
  return customers;
}

/**
 * Save a new customer or update an existing one in the 'Customers' sheet.
 */
function saveCustomer(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Customers");

  if (!sheet) {
    sheet = ss.insertSheet("Customers");
    sheet.appendRow(["Customer Name"]);
  }

  const customerName = data.name.trim();

  if (data.rowIndex) {
    // Update existing row
    sheet.getRange(data.rowIndex, 1, 1, 1).setValue(customerName);
  } else {
    // Add new row
    sheet.appendRow([customerName]);
  }

  return { success: true };
}

/**
 * Delete a customer row from the 'Customers' sheet.
 */
function deleteCustomer(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  if (sheet && rowIndex) {
    sheet.deleteRow(rowIndex);
  }
  return { success: true };
}