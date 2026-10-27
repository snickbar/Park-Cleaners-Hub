/**
 * Screens the app can serve, by ?page= name. adminOnly pages are hidden
 * from (and refused to) everyone but the admin.
 */
const PAGES = {
  Index: {},
  CreateInvoice: {},
  ViewPastInvoice: {},
  ServicesManager: {},
  CustomersManager: {},
  UserAdmin: { adminOnly: true }
};

// History sheet columns (1-based)
const HISTORY_HEADERS = ["Invoice #", "Date", "Customer", "Items", "Total", "PDF Link", "Sent On", "Paid On"];
const HISTORY_COL = { invoiceNum: 1, pdfUrl: 6, sent: 7, paid: 8 };

/**
 * Escapes HTML special characters so untrusted strings can't break out of
 * markup when interpolated into generated HTML (invoice PDFs, etc.).
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* =========================================================================
   ACCESS CONTROL
   The web app runs as the person using it (executeAs USER_ACCESSING), so
   Google enforces sharing on the spreadsheet and Invoices folder, and we
   can see who they are. On top of that, only the admin (the spreadsheet
   owner) and emails on the "Users" sheet may use the app.
   ========================================================================= */

function currentUserEmail_() {
  return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

/**
 * The admin is the spreadsheet owner, unless overridden by the ADMIN_EMAIL
 * script property. Cached in script properties after the first lookup.
 */
function getAdminEmail_() {
  const props = PropertiesService.getScriptProperties();
  let admin = props.getProperty('ADMIN_EMAIL');
  if (!admin) {
    const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
    admin = owner ? owner.getEmail() : '';
    if (admin) props.setProperty('ADMIN_EMAIL', admin.toLowerCase());
  }
  return (admin || '').trim().toLowerCase();
}

function isAdmin_(email) {
  return !!email && email === getAdminEmail_();
}

function getUsersSheet_(createIfMissing) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Email", "Added On", "Added By"]);
    // Only the owner may edit this sheet, so a staff member with edit access
    // to the spreadsheet can't add people to the allowlist themselves.
    const protection = sheet.protect().setDescription("App users - admin only");
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  }
  return sheet;
}

function getAllowedEmails_() {
  const sheet = getUsersSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim().toLowerCase())
    .filter(String);
}

function isAllowedUser_(email) {
  if (!email) return false;
  return isAdmin_(email) || getAllowedEmails_().indexOf(email) !== -1;
}

/** Call at the top of every server function the browser can reach. */
function requireUser_() {
  const email = currentUserEmail_();
  if (!isAllowedUser_(email)) {
    throw new Error("Access denied. Ask the admin to add " + (email || "your Google account") + ".");
  }
  return email;
}

function requireAdmin_() {
  const email = currentUserEmail_();
  if (!isAdmin_(email)) throw new Error("Only the admin can do this.");
  return email;
}

/* =========================================================================
   PAGE SERVING
   doGet serves the first screen (?page=Name, default Index). Moving between
   screens then goes through loadPage(), which the browser swaps in with
   document.write() - this keeps working inside the Google Sites embed,
   where links aren't allowed to replace the top-level page. See the note
   in Common.html on why page scripts declare their globals with `var`.
   ========================================================================= */

function doGet(e) {
  let email = '';
  try {
    email = currentUserEmail_();
    if (!getAdminEmail_()) {
      return accessDeniedPage_(email, "No admin is set up yet. In the Apps Script editor, open Project Settings > " +
        "Script Properties and add ADMIN_EMAIL with the admin's Google account email.");
    }
    if (!isAllowedUser_(email)) return accessDeniedPage_(email);

    // First admin visit: remember which Drive folder holds the invoices, so
    // staff (whose "Invoices" search would find their own folders) use it too.
    if (isAdmin_(email)) getInvoiceFolder_();

    const output = renderPage_((e && e.parameter && e.parameter.page) || 'Index', email);
    output.setTitle('Park Cleaners Invoice Manager');
    output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return output;
  } catch (err) {
    // Typically: the spreadsheet isn't shared with this account at all.
    console.error("doGet failed for " + email + ": " + (err && err.stack || err));
    // Show the admin where it failed; everyone else just gets the message
    let showStack = false;
    try { showStack = !!email && email === PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL'); } catch (e) {}
    return accessDeniedPage_(email, "Error: " + (err && (showStack ? err.stack : err.message) || err));
  }
}

/** Returns a screen's HTML for in-place navigation (see Common.html goTo). */
function loadPage(pageName) {
  const email = requireUser_();
  return renderPage_(pageName, email).getContent();
}

/** Evaluates a screen's template. Unknown or forbidden pages fall back to Index. */
function renderPage_(pageName, email) {
  const admin = isAdmin_(email);
  if (!PAGES.hasOwnProperty(pageName) || (PAGES[pageName].adminOnly && !admin)) {
    pageName = 'Index';
  }

  const baseUrl = ScriptApp.getService().getUrl();
  const common = HtmlService.createTemplateFromFile('Common');
  common.baseUrl = baseUrl;

  const template = HtmlService.createTemplateFromFile(pageName);
  template.commonHtml = common.evaluate().getContent();
  template.baseUrl = baseUrl;
  template.isAdmin = admin;
  template.userEmail = email;
  return template.evaluate();
}

function accessDeniedPage_(email, details) {
  const who = email ? escapeHtml_(email) : "your Google account";
  return HtmlService.createHtmlOutput(
    '<div style="font-family: Arial, sans-serif; max-width: 560px; margin: 80px auto; text-align: center;">' +
    '<h2 style="color: #c62828;">No access</h2>' +
    '<p style="font-size: 18px;">' + who + ' is not allowed to use the Park Cleaners Invoice Manager.</p>' +
    '<p style="font-size: 16px; color: #555;">Ask the admin to add you, then reload this page.</p>' +
    (details ? '<p style="font-size: 13px; color: #888;">' + escapeHtml_(details) + '</p>' : '') +
    '</div>'
  ).setTitle('Park Cleaners Invoice Manager');
}

/* =========================================================================
   DRIVE FOLDERS
   ========================================================================= */

function getInvoiceFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('INVOICE_FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);

  if (!isAdmin_(currentUserEmail_())) {
    throw new Error("The Invoices folder isn't set up yet. Ask the admin to open the app once.");
  }
  const folders = DriveApp.getFoldersByName("Invoices");
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Invoices");
  props.setProperty('INVOICE_FOLDER_ID', folder.getId());
  return folder;
}

function getBulkPrintFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BULK_PRINT_FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);

  const folder = getInvoiceFolder_().createFolder("Bulk Print");
  props.setProperty('BULK_PRINT_FOLDER_ID', folder.getId());
  return folder;
}

/* =========================================================================
   CUSTOMERS & SERVICES (read)
   ========================================================================= */

/**
 * Fetch customer list from the 'Customers' sheet.
 */
function getCustomerList() {
  requireUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return [];

  // Display values: see getInvoiceHistory on why Dates must stay out of replies
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(row => String(row[0]).trim()).filter(String);
}

/**
 * Fetch service list and prices from the 'Services' sheet.
 */
function getServiceList() {
  requireUser_();
  return readServiceList_();
}

function readServiceList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Services");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(row => {
    return {
      service: String(row[0]).trim(),
      price: Number(row[1]) || 0
    };
  }).filter(item => item.service);
}

/* =========================================================================
   INVOICES
   ========================================================================= */

function getHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  if (!sheet) {
    sheet = ss.insertSheet("History");
    sheet.appendRow(HISTORY_HEADERS);
  } else if (sheet.getLastColumn() < HISTORY_HEADERS.length) {
    // Older sheets predate the Sent On / Paid On columns. The sheet may also
    // have had its spare columns deleted, so make room before writing.
    const missing = HISTORY_HEADERS.length - sheet.getMaxColumns();
    if (missing > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), missing);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  }
  return sheet;
}

/** Returns the 1-based History row for an invoice number, or 0. */
function findHistoryRow_(sheet, invoiceNum) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const nums = sheet.getRange(2, HISTORY_COL.invoiceNum, lastRow - 1, 1).getValues();
  const target = String(invoiceNum).trim();
  for (let i = 0; i < nums.length; i++) {
    if (String(nums[i][0]).trim() === target) return i + 2;
  }
  return 0;
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

/**
 * Re-prices line items from the Services sheet instead of trusting the
 * prices and totals sent by the browser.
 */
function normalizeItems_(rawItems) {
  if (!Array.isArray(rawItems) || rawItems.length === 0) {
    throw new Error("An invoice needs at least one line item.");
  }
  const prices = {};
  readServiceList_().forEach(s => prices[s.service] = s.price);

  return rawItems.map(raw => {
    const service = String(raw.service || '').trim();
    if (!prices.hasOwnProperty(service)) {
      throw new Error(`"${service}" is no longer in the services list. Remove it from the invoice and add it again.`);
    }
    const quantity = Number(raw.quantity);
    if (!isFinite(quantity) || quantity <= 0) {
      throw new Error(`Invalid quantity for "${service}".`);
    }
    const price = prices[service];
    return {
      service: service,
      notes: String(raw.notes || '').trim().slice(0, 500),
      quantity: quantity,
      price: price,
      total: round2_(price * quantity)
    };
  });
}

/**
 * Splits items into pages by estimated printed height rather than a flat row
 * count, since long descriptions/notes wrap onto several lines.
 * Budgets are in "single-line row" units; tune after checking real PDFs.
 */
function paginateItems_(items) {
  const LINES_PER_PAGE = 16;
  const DESC_CHARS_PER_LINE = 36;
  const NOTES_CHARS_PER_LINE = 26;

  const cost = item => {
    const lines = Math.max(
      1,
      Math.ceil(item.service.length / DESC_CHARS_PER_LINE),
      Math.ceil((item.notes || '').length / NOTES_CHARS_PER_LINE)
    );
    return 1 + (lines - 1) * 0.55; // extra wrapped lines are shorter than a full row
  };

  const pages = [[]];
  let used = 0;
  items.forEach(item => {
    const c = cost(item);
    if (used + c > LINES_PER_PAGE && pages[pages.length - 1].length > 0) {
      pages.push([]);
      used = 0;
    }
    pages[pages.length - 1].push(item);
    used += c;
  });
  return pages;
}

/**
 * Finalizes the invoice, logs to History, builds a real PDF matching exact format,
 * and saves it into the "Invoices" Drive folder.
 */
function finalizeInvoice(invoiceData) {
  requireUser_();

  const customer = String(invoiceData.customer || '').trim();
  if (!customer) throw new Error("Please select a customer.");
  const items = normalizeItems_(invoiceData.items);
  const total = round2_(items.reduce((sum, item) => sum + item.total, 0));

  // Format dates (MM/DD/YYYY)
  const rawDate = invoiceData.date ? new Date(invoiceData.date + "T00:00:00") : new Date();
  const tz = Session.getScriptTimeZone();
  const serviceDateFormatted = Utilities.formatDate(rawDate, tz, "MM/dd/yyyy");
  const dateIssuedFormatted = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");

  const invoiceFolder = getInvoiceFolder_();
  const historySheet = getHistorySheet_();

  const itemsSummary = items.map(item => {
    const noteText = item.notes ? ` (${item.notes})` : '';
    return `${item.quantity}x ${item.service}${noteText} [$${item.price.toFixed(2)}]`;
  }).join("; ");

  // Reserve the invoice number by writing its History row while holding a
  // lock, so two people finalizing at once can't get the same number. The
  // number is based on the highest existing one, or the highest deleted one
  // (see deleteInvoice), so a deleted invoice's number is never reused.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let invoiceNum;
  try {
    let maxInvoiceNum = Math.max(1079, highestDeletedInvoiceNum_());
    const lastRow = historySheet.getLastRow();
    if (lastRow > 1) {
      historySheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(row => {
        const n = Number(row[0]);
        if (n > maxInvoiceNum) maxInvoiceNum = n;
      });
    }
    invoiceNum = maxInvoiceNum + 1;
    historySheet.appendRow([invoiceNum, dateIssuedFormatted, customer, itemsSummary, total, '', '', '']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  let pdfFile = null;
  try {
    pdfFile = createInvoicePdf_({
      invoiceNum: invoiceNum,
      customer: customer,
      items: items,
      total: total,
      serviceDateFormatted: serviceDateFormatted,
      dateIssuedFormatted: dateIssuedFormatted
    }, invoiceFolder);
    const pdfUrl = pdfFile.getUrl();
    historySheet.getRange(findHistoryRow_(historySheet, invoiceNum), HISTORY_COL.pdfUrl).setValue(pdfUrl);

    return {
      success: true,
      invoiceNum: invoiceNum,
      pdfUrl: pdfUrl
    };
  } catch (err) {
    // The invoice wasn't recorded, so release the reserved number rather than
    // leave a History row that points nowhere, and bin any PDF already made
    // so it can't be mixed up with the next invoice to get that number.
    const row = findHistoryRow_(historySheet, invoiceNum);
    if (row) historySheet.deleteRow(row);
    if (pdfFile) {
      try { pdfFile.setTrashed(true); } catch (e) { /* best effort */ }
    }
    throw err;
  }
}

function createInvoicePdf_(inv, invoiceFolder) {
  const pages = paginateItems_(inv.items);
  const totalPages = pages.length;

  const headerHtml = `
    <div class="company-header">
      <div class="company-name">Park Cleaners</div>
      <div class="company-sub">
        Professional Cleaning &amp; Tailoring Services<br>
        97 Pehle Ave, Saddle Brook, NJ 07663<br>
        201-388-1044 &bull; park.cleaners@yahoo.com
      </div>
    </div>

    <table class="meta-container">
      <tr>
        <td style="vertical-align: top;">
          <div class="bill-to-label">Bill To</div>
          <div class="bill-to-name">${escapeHtml_(inv.customer)}</div>
        </td>
        <td style="vertical-align: top;" class="meta-right">
          <div>Invoice #: ${inv.invoiceNum}</div>
          <div style="margin-top: 3px;">Date Issued: ${inv.dateIssuedFormatted}</div>
        </td>
      </tr>
    </table>

    <hr class="divider">
  `;

  const continuationHeaderHtml = `
    <div class="continuation-header">
      Park Cleaners &mdash; Invoice #${inv.invoiceNum} for ${escapeHtml_(inv.customer)} (continued)
    </div>
  `;

  let pagesHtml = "";
  pages.forEach(function(pageItems, p) {
    const isLastPage = p === totalPages - 1;

    let rowsHtml = "";
    pageItems.forEach(function(item) {
      rowsHtml += `
        <tr>
          <td style="border: 1px solid #000000; padding: 5px 8px; font-weight: bold;">${inv.serviceDateFormatted}</td>
          <td style="border: 1px solid #000000; padding: 5px 8px; font-weight: bold;">${escapeHtml_(item.service)}</td>
          <td style="border: 1px solid #000000; padding: 5px 8px; font-weight: bold;">${escapeHtml_(item.notes || '')}</td>
          <td style="border: 1px solid #000000; padding: 5px 8px; font-weight: bold; text-align: left;">${item.quantity}</td>
          <td style="border: 1px solid #000000; padding: 5px 8px; font-weight: bold;">$${item.price.toFixed(2)}</td>
          <td style="border: 1px solid #000000; padding: 5px 8px; font-weight: bold;">$${item.total.toFixed(2)}</td>
        </tr>
      `;
    });

    pagesHtml += `
      <div class="invoice-page${isLastPage ? '' : ' page-break'}">
        ${p === 0 ? headerHtml : continuationHeaderHtml}

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
            ${rowsHtml}
          </tbody>
        </table>

        ${isLastPage ? `
          <div class="total-container">
            TOTAL: $${inv.total.toFixed(2)}
          </div>
          <div class="footer-container">
            Thank you for your business!<br>
            Barry Barone
          </div>
        ` : ''}

        <div class="page-footer">
          Invoice #${inv.invoiceNum} &bull; Date Issued: ${inv.dateIssuedFormatted} &bull; Page ${p + 1} of ${totalPages}
        </div>
      </div>
    `;
  });

  // Build full printable HTML layout
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
          margin: 0;
          color: #000000;
          background: #ffffff;
        }

        /* Each .invoice-page holds one physical page's worth of rows;
           page-break forces a new page before the next .invoice-page. */
        .page-break {
          page-break-after: always;
        }
        .invoice-table tr {
          page-break-inside: avoid;
        }

        /* HEADER SECTION */
        .company-header {
          text-align: center;
          margin-bottom: 18px;
        }
        .company-name {
          font-size: 24px;
          font-weight: 800;
          letter-spacing: -0.5px;
          margin-bottom: 4px;
        }
        .company-sub {
          font-size: 11px;
          font-weight: bold;
          line-height: 1.3;
        }

        /* CONTINUATION HEADER (page 2+) */
        .continuation-header {
          font-size: 11px;
          font-weight: bold;
          color: #555555;
          margin-bottom: 10px;
          padding-bottom: 6px;
          border-bottom: 1px solid #cccccc;
        }

        /* METADATA SECTION */
        .meta-container {
          width: 100%;
          margin-bottom: 12px;
        }
        .bill-to-label {
          font-size: 12px;
          font-weight: bold;
        }
        .bill-to-name {
          font-size: 14px;
          font-weight: bold;
          margin-top: 3px;
        }
        .meta-right {
          text-align: right;
          font-size: 14px;
          font-weight: bold;
        }

        .divider {
          border: none;
          border-top: 1px solid #777777;
          margin-top: 8px;
          margin-bottom: 14px;
        }

        /* INVOICE TABLE */
        .invoice-table {
          width: 100%;
          border-collapse: collapse;
          border: 1px solid #000000;
        }
        .invoice-table th {
          background-color: #f2f2f2;
          border: 1px solid #000000;
          padding: 6px 8px;
          text-align: left;
          font-size: 11px;
          font-weight: 800;
          text-transform: uppercase;
        }
        .invoice-table td {
          font-size: 10.5px;
        }

        /* TOTAL SECTION */
        .total-container {
          text-align: right;
          margin-top: 18px;
          font-size: 18px;
          font-weight: 800;
          letter-spacing: 0.5px;
        }

        /* THANK-YOU FOOTER (last page only) */
        .footer-container {
          text-align: center;
          margin-top: 24px;
          font-size: 13px;
          font-weight: bold;
          line-height: 1.3;
        }

        /* PAGE FOOTER (invoice #, date, page count). Kept in normal flow
           below each page's content: a position:fixed footer per page would
           repeat on every printed page and stack on top of the others. */
        .page-footer {
          margin-top: 18px;
          text-align: center;
          font-size: 9px;
          font-weight: bold;
          color: #555555;
          border-top: 1px solid #cccccc;
          padding-top: 4px;
        }
      </style>
    </head>
    <body>
      ${pagesHtml}
    </body>
    </html>
  `;

  const blob = Utilities.newBlob(fullPdfHtml, 'text/html', `Invoice_${inv.invoiceNum}_${inv.customer}.html`);
  const pdfFile = invoiceFolder.createFile(blob.getAs('application/pdf'));
  // A public link is a convenience only: app users can already open the PDF
  // through the shared Invoices folder. Google refuses public sharing from
  // some accounts ("Limit Exceeded: Drive"), and that mustn't fail the invoice.
  try {
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    console.warn("Couldn't make invoice " + inv.invoiceNum + " PDF public: " + err.message);
  }
  return pdfFile;
}

/**
 * Fetches all invoice history records and unique customer list for searching/sorting.
 */
function getInvoiceHistory() {
  requireUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("History");
  const customerSheet = ss.getSheetByName("Customers");

  let invoices = [];
  if (historySheet && historySheet.getLastRow() > 1) {
    // Use getDisplayValues() instead of getValues() to read exact formatted cell text.
    // Read only the columns the sheet has (it may predate Sent On / Paid On).
    const cols = Math.min(historySheet.getMaxColumns(), HISTORY_HEADERS.length);
    const data = historySheet
      .getRange(2, 1, historySheet.getLastRow() - 1, cols)
      .getDisplayValues()
      .map(row => row.concat(new Array(HISTORY_HEADERS.length - cols).fill('')));
    invoices = data.map(row => {
      return {
        invoiceNum: row[0],
        date: row[1],
        customer: row[2],
        items: row[3],
        // Strip any currency formatting so the page can add its own "$"
        total: row[4].replace(/[^0-9.-]/g, ''),
        pdfUrl: row[5],
        sent: row[6],
        paid: row[7]
      };
    }).filter(inv => inv.invoiceNum);
  }

  let customers = [];
  if (customerSheet) {
    // Display values, so a cell Sheets reads as a date can't put a Date in the
    // reply (google.script.run turns a reply containing a Date into null)
    const custData = customerSheet.getDataRange().getDisplayValues();
    if (custData.length > 1) {
      customers = custData.slice(1).map(r => String(r[0]).trim()).filter(String);
    }
  }

  return {
    invoices: invoices,
    customers: customers
  };
}

/**
 * Marks an invoice as sent or paid (done = true), or clears the mark.
 * Returns the stored date text ('' when cleared).
 */
function setInvoiceStatus(invoiceNum, field, done) {
  requireUser_();
  const col = HISTORY_COL[field];
  if (field !== 'sent' && field !== 'paid') throw new Error("Unknown status: " + field);

  const sheet = getHistorySheet_();
  const row = findHistoryRow_(sheet, invoiceNum);
  if (!row) throw new Error("Invoice #" + invoiceNum + " was not found.");

  const cell = sheet.getRange(row, col);
  cell.setValue(done ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy") : '');
  SpreadsheetApp.flush();
  return cell.getDisplayValue();
}

/**
 * Marks several invoices as sent or paid with today's date. Invoices that
 * already have that mark keep their original date. Returns
 * { invoiceNum: stored date text } for every invoice found.
 */
function setInvoiceStatusBulk(invoiceNums, field) {
  requireUser_();
  if (field !== 'sent' && field !== 'paid') throw new Error("Unknown status: " + field);
  const col = HISTORY_COL[field];

  const sheet = getHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const nums = sheet.getRange(2, HISTORY_COL.invoiceNum, lastRow - 1, 1).getValues();
  const statusRange = sheet.getRange(2, col, lastRow - 1, 1);
  const current = statusRange.getDisplayValues();

  const wanted = new Set((invoiceNums || []).map(n => String(n).trim()));
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  const rows = [];
  nums.forEach((row, i) => {
    if (wanted.has(String(row[0]).trim())) rows.push(i);
  });

  // Write only the blank cells, one by one, so a status someone else
  // changes at the same moment on another row isn't overwritten.
  rows.forEach(i => {
    if (!current[i][0]) sheet.getRange(i + 2, col).setValue(today);
  });
  SpreadsheetApp.flush();

  const stored = statusRange.getDisplayValues();
  const result = {};
  rows.forEach(i => { result[String(nums[i][0]).trim()] = stored[i][0]; });
  return result;
}

function highestDeletedInvoiceNum_() {
  return Number(PropertiesService.getScriptProperties().getProperty('HIGHEST_DELETED_INVOICE')) || 0;
}

/**
 * Removes an invoice from History after moving its PDF to the Drive trash.
 * Its number is remembered so finalizeInvoice never hands it out again.
 * Only a PDF's owner can trash it, so if that's refused nothing is deleted.
 */
function deleteInvoice(invoiceNum) {
  requireUser_();
  const sheet = getHistorySheet_();

  // Same lock as finalizeInvoice, so a new invoice can't be numbered
  // between the row going and its number being recorded.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const row = findHistoryRow_(sheet, invoiceNum);
    if (!row) throw new Error("Invoice #" + invoiceNum + " was not found.");

    const fileId = driveIdFromUrl_(sheet.getRange(row, HISTORY_COL.pdfUrl).getValue());
    let file = null;
    if (fileId) {
      try {
        file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
      } catch (err) {
        let owner = '';
        try { owner = file && file.getOwner() ? file.getOwner().getEmail() : ''; } catch (e) {}
        throw new Error("Its PDF couldn't be moved to the trash" +
          (owner ? ", because it belongs to " + owner + ". Ask them to delete this invoice." : ": " + err.message));
      }
    }

    try {
      const num = Number(invoiceNum);
      if (num > highestDeletedInvoiceNum_()) {
        PropertiesService.getScriptProperties().setProperty('HIGHEST_DELETED_INVOICE', String(num));
      }
      sheet.deleteRow(row);
      SpreadsheetApp.flush();
    } catch (err) {
      if (file) file.setTrashed(false); // keep the PDF with its invoice
      throw err;
    }
  } finally {
    lock.releaseLock();
  }
}

function driveIdFromUrl_(url) {
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

/**
 * Returns the PDFs for up to 10 invoices as base64, for merging into one
 * printable file in the browser. Looked up by invoice number so the page
 * can't ask for arbitrary Drive files.
 */
function getInvoicePdfs(invoiceNums) {
  requireUser_();
  if (!Array.isArray(invoiceNums) || invoiceNums.length > 10) {
    throw new Error("Request at most 10 invoices at a time.");
  }
  const sheet = getHistorySheet_();
  return invoiceNums.map(num => {
    const row = findHistoryRow_(sheet, num);
    const fileId = row ? driveIdFromUrl_(sheet.getRange(row, HISTORY_COL.pdfUrl).getValue()) : '';
    if (!fileId) throw new Error("Invoice #" + num + " has no PDF.");
    const bytes = DriveApp.getFileById(fileId).getBlob().getBytes();
    return { invoiceNum: String(num), data: Utilities.base64Encode(bytes) };
  });
}

/**
 * Saves a merged bulk-print PDF to Invoices/Bulk Print and returns its link.
 * Bulk files older than a day are cleaned up as new ones are made.
 */
function saveBulkPrintPdf(base64Pdf, invoiceCount) {
  requireUser_();
  const folder = getBulkPrintFolder_();

  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  const old = folder.getFiles();
  while (old.hasNext()) {
    const f = old.next();
    if (f.getDateCreated().getTime() < cutoff) {
      try { f.setTrashed(true); } catch (e) { /* made by someone else; leave it */ }
    }
  }

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Pdf),
    'application/pdf',
    `Bulk_Print_${Number(invoiceCount) || 0}_invoices_${stamp}.pdf`
  );
  return folder.createFile(blob).getUrl();
}

/* =========================================================================
   SERVICES MANAGER
   ========================================================================= */

/**
 * Throws unless rowIndex still holds expectedName, so an edit or delete made
 * from a stale page can't hit the wrong row after the sheet has changed.
 */
function assertRowMatches_(sheet, rowIndex, expectedName, listName) {
  const row = Number(rowIndex);
  const ok = Number.isInteger(row) && row >= 2 && row <= sheet.getLastRow() &&
    sheet.getRange(row, 1).getDisplayValue().trim() === String(expectedName || '').trim();
  if (!ok) {
    throw new Error(`The ${listName} list changed since this page loaded. It has been reloaded - please try again.`);
  }
}

/**
 * Fetch all services from the 'Services' sheet for the Services Manager page.
 */
function getServicesData() {
  requireUser_();
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
  requireUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Services");

  if (!sheet) {
    sheet = ss.insertSheet("Services");
    sheet.appendRow(["Service", "Price"]);
  }

  const serviceName = String(data.service || '').trim();
  if (!serviceName) throw new Error("Service name is required.");
  const priceVal = parseFloat(data.price) || 0;
  const formattedPrice = priceVal.toFixed(2);

  if (data.rowIndex) {
    // Update existing row
    assertRowMatches_(sheet, data.rowIndex, data.originalService, "services");
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
function deleteService(rowIndex, expectedService) {
  requireUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Services");
  if (sheet && rowIndex) {
    assertRowMatches_(sheet, rowIndex, expectedService, "services");
    sheet.deleteRow(rowIndex);
  }
  return { success: true };
}

/* =========================================================================
   CUSTOMERS MANAGER
   ========================================================================= */

/**
 * Fetch all customers from the 'Customers' sheet for the Customers Manager page.
 */
function getCustomersData() {
  requireUser_();
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
  requireUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Customers");

  if (!sheet) {
    sheet = ss.insertSheet("Customers");
    sheet.appendRow(["Customer Name"]);
  }

  const customerName = String(data.name || '').trim();
  if (!customerName) throw new Error("Customer name is required.");

  if (data.rowIndex) {
    // Update existing row
    assertRowMatches_(sheet, data.rowIndex, data.originalName, "customers");
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
function deleteCustomer(rowIndex, expectedName) {
  requireUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  if (sheet && rowIndex) {
    assertRowMatches_(sheet, rowIndex, expectedName, "customers");
    sheet.deleteRow(rowIndex);
  }
  return { success: true };
}

/* =========================================================================
   USER ADMIN (admin only)
   Adding a user shares the spreadsheet and Invoices folder with them (they
   need both, since the app runs as them) and puts them on the allowlist.
   ========================================================================= */

function getUsers() {
  requireAdmin_();
  const sheet = getUsersSheet_(false);
  let users = [];
  if (sheet && sheet.getLastRow() > 1) {
    users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues()
      .map(r => ({ email: r[0].trim().toLowerCase(), addedOn: r[1] }))
      .filter(u => u.email);
  }
  return {
    adminEmail: getAdminEmail_(),
    appUrl: ScriptApp.getService().getUrl(),
    users: users
  };
}

function addUser(rawEmail) {
  const adminEmail = requireAdmin_();
  const email = String(rawEmail || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("That doesn't look like an email address.");
  if (email === adminEmail) throw new Error("You're the admin - you always have access.");
  if (getAllowedEmails_().indexOf(email) !== -1) throw new Error(email + " already has access.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = getInvoiceFolder_();
  ss.addEditor(email);
  folder.addEditor(email);
  // Stop staff from re-sharing the data with people the admin didn't add
  DriveApp.getFileById(ss.getId()).setShareableByEditors(false);
  folder.setShareableByEditors(false);

  const tz = Session.getScriptTimeZone();
  getUsersSheet_(true).appendRow([email, Utilities.formatDate(new Date(), tz, "MM/dd/yyyy"), adminEmail]);
  return getUsers();
}

function removeUser(rawEmail) {
  requireAdmin_();
  const email = String(rawEmail || '').trim().toLowerCase();
  const sheet = getUsersSheet_(false);
  if (sheet && sheet.getLastRow() > 1) {
    const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    // Bottom-up so deleting a row doesn't shift ones not yet checked
    for (let i = emails.length - 1; i >= 0; i--) {
      if (String(emails[i][0]).trim().toLowerCase() === email) sheet.deleteRow(i + 2);
    }
  }
  try { SpreadsheetApp.getActiveSpreadsheet().removeEditor(email); } catch (e) { /* not an editor */ }
  try { getInvoiceFolder_().removeEditor(email); } catch (e) { /* not an editor */ }
  return getUsers();
}
