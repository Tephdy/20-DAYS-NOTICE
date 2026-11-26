function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'load') return loadDatabase();
    if (data.action === 'save' || data.action === 'update') return handleCRUD(data);
    if (data.action === 'clearTab') return handleClearTab(data);

    const fileName = data.fileName;
    const base64Data = data.base64;
    const rootFolderId = data.rootFolderId;
    const propertyName = data.propertyName || "Unknown Property";
    let monthName = data.monthName || data.folderMonth || "Unknown Month";
    const tenantName = data.tenantName || "Unknown Tenant";

    const rootFolder = DriveApp.getFolderById(rootFolderId);

    let propertyFolder;
    const propertyFolders = rootFolder.getFoldersByName(propertyName);
    if (propertyFolders.hasNext()) propertyFolder = propertyFolders.next();
    else propertyFolder = rootFolder.createFolder(propertyName);

    let monthFolder;
    const monthFolders = propertyFolder.getFoldersByName(monthName);
    if (monthFolders.hasNext()) monthFolder = monthFolders.next();
    else monthFolder = propertyFolder.createFolder(monthName);

    let tenantFolder;
    const tenantFolders = monthFolder.getFoldersByName(tenantName);
    if (tenantFolders.hasNext()) tenantFolder = tenantFolders.next();
    else tenantFolder = monthFolder.createFolder(tenantName);

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', fileName);
    const file = tenantFolder.createFile(blob);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      fileUrl: file.getUrl()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

const SPREADSHEET_ID = "1Gmbssfz0_72riquWeVEAlDli6d6zM_ET";
const STANDARD_HEADERS = [
  "Tenant Name", "Unit", "Building", "Computed Due Date", "Monthly Rent",
  "Status", "Date Sent", "PDF Link", "Term", "Reason", "Start Date", "End of Contract"
];

function calculateDynamicDate(startDateStr, termStr) {
  if (!startDateStr) return "";
  const startDate = new Date(startDateStr);
  if (isNaN(startDate.getTime())) return startDateStr;

  const termString = String(termStr).toUpperCase().trim();
  const isMonthly = termString === 'M-M' || termString === 'MM' ||
    termString.includes('MONTHLY') || termString === 'TRIAL MONTH' ||
    termString === 'EXT. TRIAL' || termString === 'EXT. TRIAL +' ||
    termString === 'MONTH-MONTH';
  const termMonths = isMonthly ? 1 : (Number(termStr) || 12);

  const actualEndDate = new Date(startDate);
  actualEndDate.setMonth(actualEndDate.getMonth() + termMonths);

  const noticeDate = new Date(actualEndDate);
  noticeDate.setDate(noticeDate.getDate() - (isMonthly ? 10 : 20));

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  while (noticeDate.getTime() < today.getTime() && noticeDate.getTime() < actualEndDate.getTime()) {
    noticeDate.setMonth(noticeDate.getMonth() + 1);
  }
  if (noticeDate.getTime() > actualEndDate.getTime()) {
    noticeDate.setTime(actualEndDate.getTime());
  }

  const year = noticeDate.getFullYear();
  const month = String(noticeDate.getMonth() + 1).padStart(2, '0');
  const day = String(noticeDate.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function headerIndex(headers, headerName) {
  const expected = String(headerName).trim().toLowerCase();
  return headers.findIndex(header => String(header || '').trim().toLowerCase() === expected);
}

function normalizeEndContract(value) {
  if (value === undefined || value === null || value === '') return '';

  const dateText = String(value).trim();
  const isoMatch = dateText.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
      return dateText;
    }
    throw new Error("End of Contract must be a valid calendar date.");
  }

  return dateText;
}

function ensureEndContractColumn(sheet, headers) {
  if (headerIndex(headers, "End of Contract") !== -1) return headers;

  const endColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, endColumn).setValue("End of Contract");
  headers.push("End of Contract");
  return headers;
}

function loadDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const allData = {};

  sheets.forEach(sheet => {
    const tabName = sheet.getName();
    if (tabName === "System Exports") return;

    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return;

    let lastCol = sheet.getLastColumn();
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    if (headers[0] !== "Tenant Name" || headers[6] !== "Date Sent" || headers.length < 11) {
      sheet.getRange(1, 1, 1, STANDARD_HEADERS.length).setValues([STANDARD_HEADERS]);
      headers = STANDARD_HEADERS.slice();
      lastCol = STANDARD_HEADERS.length;
    } else {
      headers = ensureEndContractColumn(sheet, headers);
      lastCol = sheet.getLastColumn();
    }

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const richText = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();
    headers = data[0];
    const pdfColIndex = headerIndex(headers, "PDF Link");
    const computedCol = headerIndex(headers, "Computed Due Date");

    const rows = data.slice(1).map((row, i) => {
      let pdfUrl = row[pdfColIndex] || '';
      if (richText[i + 1] && richText[i + 1][pdfColIndex] && richText[i + 1][pdfColIndex].getLinkUrl()) {
        pdfUrl = richText[i + 1][pdfColIndex].getLinkUrl();
      }

      const getVal = headerName => {
        const index = headerIndex(headers, headerName);
        return index === -1 ? '' : row[index] || '';
      };

      const startDateStr = getVal("Start Date");
      const dynamicDate = calculateDynamicDate(startDateStr, getVal("Term"));

      if (dynamicDate && computedCol !== -1 && dynamicDate !== row[computedCol]) {
        sheet.getRange(i + 2, computedCol + 1).setValue(dynamicDate);
      }

      return {
        tenantName: getVal("Tenant Name"),
        unit: getVal("Unit"),
        building: getVal("Building"),
        computed20th: dynamicDate || getVal("Computed Due Date"),
        rent: getVal("Monthly Rent") || 0,
        status: getVal("Status") || 'Pending',
        pdfUrl: pdfUrl,
        termLabel: getVal("Term") || '1 year',
        reason: getVal("Reason") || '—',
        dateSent: getVal("Date Sent") || '',
        startDate: startDateStr,
        endContract: getVal("End of Contract")
      };
    });

    if (rows.length > 0) {
      allData[tabName] = rows;
    }
  });

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    data: allData
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleCRUD(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(data.tabName);

  if (!sheet) {
    sheet = ss.insertSheet(data.tabName);
    sheet.getRange(1, 1, 1, STANDARD_HEADERS.length).setValues([STANDARD_HEADERS]);
  }

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers = ensureEndContractColumn(sheet, headers);

  if (data.action === 'save') {
    const tenant = {
      "Tenant Name": data.tenantName,
      "Unit": data.unit,
      "Building": data.building,
      "Computed Due Date": data.computed20th,
      "Monthly Rent": data.rent,
      "Status": data.status,
      "Date Sent": data.dateSent || '',
      "PDF Link": data.pdfUrl || '',
      "Term": data.termLabel || '',
      "Reason": data.reason || '',
      "Start Date": data.startDate || '',
      "End of Contract": normalizeEndContract(data.endContract)
    };
    sheet.appendRow(headers.map(header => {
      const key = Object.keys(tenant).find(name => name.toLowerCase() === String(header || '').trim().toLowerCase());
      return key ? tenant[key] : '';
    }));
  } else if (data.action === 'update') {
    const tenantNameCol = headerIndex(headers, "Tenant Name") + 1;
    const unitCol = headerIndex(headers, "Unit") + 1;
    let updatedRows = 0;
    if (tenantNameCol > 0 && unitCol > 0 && sheet.getLastRow() > 1) {
      const textFinder = sheet.getRange(2, tenantNameCol, sheet.getLastRow() - 1, 1)
        .createTextFinder(data.previousTenantName || data.tenantName)
        .matchEntireCell(true)
        .findAll();

      for (let i = 0; i < textFinder.length; i++) {
        const row = textFinder[i].getRow();
        const rowUnit = sheet.getRange(row, unitCol).getValue();
        const expectedUnit = data.previousUnit !== undefined ? data.previousUnit : data.unit;

        if (String(rowUnit).trim() === String(expectedUnit).trim()) {
          const setIfPresent = (headerName, value) => {
            const column = headerIndex(headers, headerName) + 1;
            if (column > 0 && value !== undefined) sheet.getRange(row, column).setValue(value);
          };

          setIfPresent("Tenant Name", data.tenantName);
          setIfPresent("Unit", data.unit);
          setIfPresent("Building", data.building);
          setIfPresent("Computed Due Date", data.computed20th);
          setIfPresent("Monthly Rent", data.rent);
          setIfPresent("Status", data.status);
          setIfPresent("Date Sent", data.dateSent);
          setIfPresent("Term", data.termLabel);
          setIfPresent("Reason", data.reason);
          setIfPresent("Start Date", data.startDate);
          setIfPresent("End of Contract", normalizeEndContract(data.endContract));
          updatedRows++;

          if (data.pdfUrl !== undefined) {
            const pdfColumn = headerIndex(headers, "PDF Link") + 1;
            if (pdfColumn > 0) {
              const cell = sheet.getRange(row, pdfColumn);
              const url = data.pdfUrl;
              if (url && url.trim() !== '') {
                cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build());
              } else {
                cell.clearContent();
              }
            }
          }
        }
      }
    }
    if (updatedRows === 0) {
      throw new Error(`Could not find "${data.previousTenantName || data.tenantName}" in "${data.tabName}" to update.`);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
}

function handleClearTab(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.tabName);

  if (sheet) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, STANDARD_HEADERS.length).setValues([STANDARD_HEADERS]);
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
}
