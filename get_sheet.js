const { google } = require("googleapis");
const path = require("path");

// ===== Google Form Data =====
const SERVICE_ACCOUNT_FILE = path.join(__dirname, "service-account.json");
const SPREADSHEET_ID = "1iCszLbkhZOpQfV4fubIHXIp8q6ajWhE31PzgeZvFA3g";
const SHEET_NAME = "Form Responses 1";
const RANGE = `${SHEET_NAME}!A:I`;

// ==== Sheet Master Data =====
const MASTER_SPREADSHEET_ID = "1LRg_0DeuHgIwax7FV0pCYYr9NyPb9ToGxXGcS0TTDPY";
const MASTER_SHEET_NAME = "Sheet1";
const MASTER_RANGE = `${MASTER_SHEET_NAME}!A:H`;

function clean(value) {
  return (value || "").toString().trim();
}

function formatTodayMMDDYYYY() {
  const now = new Date();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  const year = now.getFullYear();
  return `${month}/${day}/${year}`;
}

async function getGoogleClients() {
  const auth = new google.auth.GoogleAuth({
    keyFile: SERVICE_ACCOUNT_FILE,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const client = await auth.getClient();
  const sheets = google.sheets({ version: "v4", auth: client });

  return { sheets };
}

function columnToLetter(column) {
  let temp = "";
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }

  return letter;
}

async function getHeaderMap() {
  const { sheets } = await getGoogleClients();

  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!1:1`,
  });

  const headers = headerRes.data.values?.[0] || [];
  const headerMap = {};

  headers.forEach((header, index) => {
    headerMap[header] = index + 1;
  });

  return { sheets, headers, headerMap };
}

async function getSheetDataTodayOnly() {
  const { sheets } = await getGoogleClients();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RANGE,
  });

  const rows = res.data.values || [];
  if (rows.length === 0) return [];

  const headers = rows[0];
  const todayText = formatTodayMMDDYYYY();

  const data = rows.slice(1).map((row, index) => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] || "";
    });

    // Store Row Id
    obj.__rowNumber = index + 2;

    return obj;
  });

  return data.filter((item) => {
    const rawDate = (item["Timestamp"] || "").trim();
    const status = (item["Status (No input required)"] || "").trim();

    if (!rawDate) return false;

    const datePart = String(rawDate || "")
      .split(",")[0]
      .split(" ")[0]
      .trim();

    const isToday = datePart === todayText;
    const isStatusEmpty = status === "";

    return isToday && isStatusEmpty;
  });
}

async function updateCellByHeader(rowNumber, headerName, value) {
  const { sheets, headerMap } = await getHeaderMap();

  const colIndex = headerMap[headerName];
  if (!colIndex) {
    throw new Error(`Column "${headerName}" not found`);
  }

  const colLetter = columnToLetter(colIndex);
  const updateRange = `${SHEET_NAME}!${colLetter}${rowNumber}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: updateRange,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[value]],
    },
  });
}

async function updateRowResult(rowNumber, statusText, errorMessage = "") {
  await updateCellByHeader(rowNumber, "Status (No input required)", statusText);
  await updateCellByHeader(
    rowNumber,
    "Error Message (No input required)",
    errorMessage,
  );
}

async function getMasterData() {
  const { sheets } = await getGoogleClients();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: MASTER_SPREADSHEET_ID,
    range: MASTER_RANGE,
  });

  const rows = res.data.values || [];
  if (rows.length === 0) return [];

  const headers = rows[0];

  return rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] || '';
    });
    return obj;
  });
}

function checkOwnerMatch(masterRows, fullName, vehicleReg) {
  const targetName = clean(fullName);
  const targetVehicleReg = clean(vehicleReg);

  const matchedOwnerRows = masterRows.filter((row) => {
    const thaiName = clean(row['Name-Surname (Thai)']);
    const engName = clean(row['Name Surname']);
    const vehicleType = clean(row['Vehicle Type']).toLowerCase();

    const nameMatched =
      targetName !== '' &&
      (targetName === thaiName || targetName === engName);

    const isCar = vehicleType === 'car';

    return nameMatched && isCar;
  });

  if (matchedOwnerRows.length === 0) {
    return {
      success: false,
      reason: 'Owner not match',
    };
  }

  const vehicleMatched = matchedOwnerRows.some((row) => {
    return clean(row['Vehicle Registration']) === targetVehicleReg;
  });

  if (!vehicleMatched) {
    return {
      success: false,
      reason: 'Owner not match',
    };
  }

  return {
    success: true,
    reason: '',
  };
}

module.exports = {
  getSheetDataTodayOnly,
  updateRowResult,
  getMasterData,
  checkOwnerMatch,
};

/*
module.exports = {
  getSheetDataTodayOnly,
  updateCellByHeader,
  updateRowResult,
};
*/


