/**
 * Serves the HTML interface as a standalone Web App webpage or sidebar.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CRM Web App')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SHEET_NAME = "Leads";
const CLIENT_DIRECTORY_SHEET_NAME = "Client Directory";
const CPM_MASTER_LIST_SHEET_NAME = "CPM 2024 Master List";
const HOME_HEALTH_SHEET_NAME = "Home Health";
const HOSPICE_SHEET_NAME = "Hospice";

const STATUS_COLUMN_NAME = "Status";
const LAST_CALLED_COLUMN_NAME = "Last Called Date";
const CALL_COUNT_COLUMN_NAME = "Call Count";
const RECENT_NOTES_COLUMN_NAME = "Recent Call Notes";
const HISTORY_COLUMN_NAME = "Call History Log";

// Base Column Name Constants
const DATE_COLUMN_NAME = "Date";
const COMPANY_COLUMN_NAME = "Company Name";
const DBA_COLUMN_NAME = "DBA"; 
const FACILITY_TYPE_COLUMN_NAME = "Facility Type";
const ADDRESS_COLUMN_NAME = "Address"; 
const EMAIL_COLUMN_NAME = "Email Address";
const PHONE_COLUMN_NAME = "Phone Number";

// FIELD CONSTANTS FROM SPREADSHEET ROW
const CORPORATE_COLUMN_NAME = "Corporate";
const LEAD_GENERATOR_COLUMN_NAME = "Lead Generator";
const ACCOUNT_OWNER_COLUMN_NAME = "Account Owner"; 
const ADMIN_NAME_COLUMN_NAME = "Admin Name";
const CONTACT_NAME_COLUMN_NAME = "Contact Name";
const CONTACT_TITLE_COLUMN_NAME = "Contact Title";
const CONTACT_EMAIL_COLUMN_NAME = "Contact Email";
const CONTACT_PHONE_COLUMN_NAME = "Contact Phone";
const CENTRAL_SUPPLY_COLUMN_NAME = "Central Supply";
const PATIENT_SUPPLY_COLUMN_NAME = "Patient Supply";
const PART_B_BILLING_COLUMN_NAME = "Part B Billing";
const HMO_BILLING_COLUMN_NAME = "HMO Billing";
const CITY_COLUMN_NAME = "City";
const STATE_COLUMN_NAME = "State";
const ZIP_COLUMN_NAME = "Zip";

const STATUSES = [
  "Potential Lead / Attempting Contact",
  "Engaged / Connected",
  "Disqualified / Unfit",
  "Negotiation / In Review",
  "Closed Won",
  "Closed Lost",
  "Active Client",
  "Onboarding in Progress",
  "At Risk",
  "Paused",
  "Past Client"
];

const CALL_OUTCOMES = [
  "To Call",
  "Call Follow-up",
  "Left Voicemail",
  "Declined",
  "Acct Demo",
  "In-Person",
  "Cost Comp",
  "Supply Audit"
];

const DEFAULT_STATUS = "Potential Lead / Attempting Contact";

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📞 CRM Menu')
      .addItem('Setup CRM Sheets & Dropdowns', 'setupCRM')
      .addItem('Open Logger Dashboard', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CRM Dashboard')
      .setWidth(450); 
  SpreadsheetApp.getUi().showSidebar(html);
}

function getActiveCRMUser() {
  const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return userEmail ? userEmail : "CRM Admin";
}

function getCurrentUserInfo() {
  const fullUser = getActiveCRMUser();
  const displayName = fullUser.includes('@') ? fullUser.split('@')[0] : fullUser;
  return {
    email: fullUser,
    name: displayName
  };
}

function getAllCallLogs(targetDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToScan = [
      { name: SHEET_NAME, label: "Leads" },
      { name: CLIENT_DIRECTORY_SHEET_NAME, label: "Client Directory" },
      { name: CPM_MASTER_LIST_SHEET_NAME, label: "CPM 2024 Master List" },
      { name: HOME_HEALTH_SHEET_NAME, label: "Home Health" },
      { name: HOSPICE_SHEET_NAME, label: "Hospice" }
    ];

    const logs = [];

    sheetsToScan.forEach(config => {
      const sheet = ss.getSheetByName(config.name);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getDataRange().getValues();
      if (!data || data.length <= 1) return;

      const headers = data[0].map(h => String(h).trim().toLowerCase());

      const dateIdx = headers.findIndex(h => h.includes("date") || h.includes("logged"));
      const nameIdx = headers.findIndex(h => h.includes("company") || h.includes("account") || h.includes("lead"));
      const phoneIdx = headers.findIndex(h => h.includes("phone"));
      
      let colIdx = -1;
      if (config.name === CLIENT_DIRECTORY_SHEET_NAME) {
        const sIdx = headers.indexOf("status");
        const oIdx = headers.indexOf("outcome");
        colIdx = sIdx >= 0 ? sIdx : (oIdx >= 0 ? oIdx : headers.findIndex(h => h.includes("status") || h.includes("outcome")));
      } else {
        const oIdx = headers.indexOf("outcome");
        const sIdx = headers.indexOf("status");
        colIdx = oIdx >= 0 ? oIdx : (sIdx >= 0 ? sIdx : headers.findIndex(h => h.includes("outcome") || h.includes("status")));
      }

      const notesIdx = headers.findIndex(h => h.includes("notes") || h.includes("history"));

      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        let accountName = nameIdx >= 0 ? String(row[nameIdx] || "").trim() : "";
        let phone = phoneIdx >= 0 ? String(row[phoneIdx] || "").trim() : "";
        let outcome = colIdx >= 0 ? String(row[colIdx] || "").trim() : "";
        if (!outcome) {
          outcome = (config.name === CLIENT_DIRECTORY_SHEET_NAME || config.name === SHEET_NAME) 
            ? DEFAULT_STATUS 
            : "To Call";
        }

        let rawDate = dateIdx >= 0 ? row[dateIdx] : "";

        let formattedDate = "";
        if (rawDate instanceof Date) {
          formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (rawDate) {
          const parsed = new Date(rawDate);
          if (!isNaN(parsed.getTime())) {
            formattedDate = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            formattedDate = String(rawDate).trim();
          }
        }

        let notesText = notesIdx >= 0 ? String(row[notesIdx] || "").trim() : "";

        if (notesText) {
          const lines = notesText.split('\n');
          let parsedAnyLine = false;

          lines.forEach(line => {
            const lineStr = line.trim();
            if (!lineStr) return;

            const match = lineStr.match(/\[Call Date:\s*(\d{4}-\d{2}-\d{2})\s*\|\s*([^\]]+)\]:\s*(.*)/);
            if (match) {
              parsedAnyLine = true;
              const logDate = match[1];
              const logMeta = match[2];
              const logNote = match[3];

              if (!targetDate || targetDate === logDate) {
                logs.push({
                  source: config.label,
                  callDate: logDate,
                  timeMeta: logMeta,
                  accountName: accountName || "Unnamed Account",
                  phone: phone || "-",
                  status: outcome,
                  outcome: outcome,
                  notes: logNote || lineStr,
                  row: i + 1
                });
              }
            }
          });

          if (!parsedAnyLine && formattedDate) {
            if (!targetDate || targetDate === formattedDate) {
              logs.push({
                source: config.label,
                callDate: formattedDate,
                timeMeta: formattedDate,
                accountName: accountName || "Unnamed Account",
                phone: phone || "-",
                status: outcome,
                outcome: outcome,
                notes: notesText,
                row: i + 1
              });
            }
          }
        } else if (formattedDate) {
          if (!targetDate || targetDate === formattedDate) {
            logs.push({
              source: config.label,
              callDate: formattedDate,
              timeMeta: formattedDate,
              accountName: accountName || "Unnamed Account",
              phone: phone || "-",
              status: outcome,
              outcome: outcome,
              notes: "No notes recorded",
              row: i + 1
            });
          }
        }
      }
    });

    logs.sort((a, b) => (b.callDate || "").localeCompare(a.callDate || ""));

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, logs: [], error: err.toString() };
  }
}

function getExportSchemaAndData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabConfigs = [
      { key: "records", sheetName: SHEET_NAME, label: "Records Table (Leads)" },
      { key: "clientDirectory", sheetName: CLIENT_DIRECTORY_SHEET_NAME, label: "Client Directory" },
      { key: "cpm2024MasterList", sheetName: CPM_MASTER_LIST_SHEET_NAME, label: "CPM 2024 Master List" },
      { key: "hospice", sheetName: HOSPICE_SHEET_NAME, label: "Hospice" },
      { key: "homeHealth", sheetName: HOME_HEALTH_SHEET_NAME, label: "Home Health" }
    ];

    const exportPayload = {};

    tabConfigs.forEach(tab => {
      const sheet = ss.getSheetByName(tab.sheetName);
      if (!sheet || sheet.getLastRow() < 1) {
        exportPayload[tab.key] = { label: tab.label, headers: [], rows: [] };
        return;
      }

      const data = sheet.getDataRange().getValues();
      if (!data || data.length === 0) {
        exportPayload[tab.key] = { label: tab.label, headers: [], rows: [] };
        return;
      }

      const headers = data[0].map(h => String(h).trim());
      const rows = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowObj = {};
        headers.forEach((h, idx) => {
          let val = row[idx];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
          } else if (val === null || val === undefined) {
            val = "";
          } else {
            val = String(val);
          }
          rowObj[h] = val;
        });
        rows.push(rowObj);
      }

      exportPayload[tab.key] = { label: tab.label, headers: headers, rows: rows };
    });

    return { success: true, data: exportPayload };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getCallTractionMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToScan = [SHEET_NAME, CLIENT_DIRECTORY_SHEET_NAME, CPM_MASTER_LIST_SHEET_NAME, HOME_HEALTH_SHEET_NAME, HOSPICE_SHEET_NAME];

    const dailyCounts = {};
    const weeklyCounts = {};
    const monthlyCounts = {};
    const yearlyCounts = {};

    const timestampRegex = /\[(?:Call Date:\s*(\d{4}-\d{2}-\d{2})|(\d{2})\/(\d{2})\/(\d{4}))/g;

    sheetsToScan.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getDataRange().getValues();
      if (!data || data.length <= 1) return;

      const headers = data[0].map(h => String(h).trim().toLowerCase());
      const dateColIdx = headers.findIndex(h => h.includes("date") || h.includes("logged"));
      const notesColIdx = headers.findIndex(h => h.includes("notes") || h.includes("history"));

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let rowDateObj = null;

        if (dateColIdx >= 0 && row[dateColIdx]) {
          let val = row[dateColIdx];
          if (val instanceof Date) {
            rowDateObj = val;
          } else if (typeof val === 'string' && val.trim()) {
            const parsed = new Date(val.trim());
            if (!isNaN(parsed.getTime())) rowDateObj = parsed;
          }
        }

        const recordDate = (dateObj) => {
          if (!dateObj || isNaN(dateObj.getTime())) return;
          const year = dateObj.getFullYear();
          const month = dateObj.getMonth() + 1;
          const day = dateObj.getDate();

          const dailyStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          dailyCounts[dailyStr] = (dailyCounts[dailyStr] || 0) + 1;

          const yearStr = `${year}`;
          yearlyCounts[yearStr] = (yearlyCounts[yearStr] || 0) + 1;

          const monthStr = `${year}-${String(month).padStart(2, '0')}`;
          monthlyCounts[monthStr] = (monthlyCounts[monthStr] || 0) + 1;

          const firstDayOfYear = new Date(year, 0, 1);
          const pastDaysOfYear = (dateObj - firstDayOfYear) / 86400000;
          const weekNum = Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
          const weekStr = `${year}-W${String(weekNum).padStart(2, '0')}`;
          weeklyCounts[weekStr] = (weeklyCounts[weekStr] || 0) + 1;
        };

        if (rowDateObj) recordDate(rowDateObj);

        if (notesColIdx >= 0 && row[notesColIdx]) {
          const historyLogText = String(row[notesColIdx]);
          let match;
          timestampRegex.lastIndex = 0;
          while ((match = timestampRegex.exec(historyLogText)) !== null) {
            let year, month, day;
            if (match[1]) {
              const parts = match[1].split('-');
              year = parseInt(parts[0], 10);
              month = parseInt(parts[1], 10);
              day = parseInt(parts[2], 10);
            } else {
              month = parseInt(match[2], 10);
              day = parseInt(match[3], 10);
              year = parseInt(match[4], 10);
            }
            const logDateObj = new Date(year, month - 1, day);
            if (rowDateObj && logDateObj.toDateString() === rowDateObj.toDateString()) continue;
            recordDate(logDateObj);
          }
        }
      }
    });

    return { daily: dailyCounts, weekly: weeklyCounts, monthly: monthlyCounts, yearly: yearlyCounts };
  } catch(e) {
    return { daily: {}, weekly: {}, monthly: {}, yearly: {} };
  }
}

function getDailyCallBreakdown() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToScan = [SHEET_NAME, CLIENT_DIRECTORY_SHEET_NAME, CPM_MASTER_LIST_SHEET_NAME, HOME_HEALTH_SHEET_NAME, HOSPICE_SHEET_NAME];

    const dailyMap = {};

    sheetsToScan.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim().toLowerCase());

      const dateIdx = headers.findIndex(h => h.includes("date") || h.includes("logged"));
      
      let statusIdx = -1;
      if (sheetName === CLIENT_DIRECTORY_SHEET_NAME) {
        const sIdx = headers.indexOf("status");
        const oIdx = headers.indexOf("outcome");
        statusIdx = sIdx >= 0 ? sIdx : (oIdx >= 0 ? oIdx : headers.findIndex(h => h.includes("status") || h.includes("outcome")));
      } else {
        const oIdx = headers.indexOf("outcome");
        const sIdx = headers.indexOf("status");
        statusIdx = oIdx >= 0 ? oIdx : (sIdx >= 0 ? sIdx : headers.findIndex(h => h.includes("outcome") || h.includes("status")));
      }

      const notesIdx = headers.findIndex(h => h.includes("notes") || h.includes("history"));

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let rawDate = dateIdx >= 0 ? row[dateIdx] : "";
        let statusVal = statusIdx >= 0 ? String(row[statusIdx] || "").trim() : "";
        if (!statusVal) {
          statusVal = (sheetName === CLIENT_DIRECTORY_SHEET_NAME || sheetName === SHEET_NAME) ? DEFAULT_STATUS : "To Call";
        }

        let formattedDate = "";
        if (rawDate instanceof Date) {
          formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (rawDate) {
          const parsed = new Date(rawDate);
          if (!isNaN(parsed.getTime())) {
            formattedDate = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        }

        if (formattedDate) {
          if (!dailyMap[formattedDate]) {
            dailyMap[formattedDate] = { totalCalls: 0, statuses: {} };
          }
          dailyMap[formattedDate].totalCalls += 1;
          dailyMap[formattedDate].statuses[statusVal] = (dailyMap[formattedDate].statuses[statusVal] || 0) + 1;
        }

        if (notesIdx >= 0 && row[notesIdx]) {
          const notesText = String(row[notesIdx]);
          const logRegex = /\[(?:Call Date:\s*(\d{4}-\d{2}-\d{2})|(\d{2})\/(\d{2})\/(\d{4}))/g;
          let match;
          while ((match = logRegex.exec(notesText)) !== null) {
            let logDate = "";
            if (match[1]) {
              logDate = match[1];
            } else {
              logDate = `${match[4]}-${String(match[2]).padStart(2, '0')}-${String(match[3]).padStart(2, '0')}`;
            }

            if (logDate && logDate !== formattedDate) {
              if (!dailyMap[logDate]) {
                dailyMap[logDate] = { totalCalls: 0, statuses: {} };
              }
              dailyMap[logDate].totalCalls += 1;
              dailyMap[logDate].statuses[statusVal] = (dailyMap[logDate].statuses[statusVal] || 0) + 1;
            }
          }
        }
      }
    });

    return { success: true, breakdown: dailyMap };
  } catch (err) {
    return { success: false, breakdown: {}, error: err.toString() };
  }
}

function getAllCompaniesData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      setupCRM();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { companies: [], statuses: STATUSES, outcomes: CALL_OUTCOMES };
    }

    const lastCol = sheet.getLastColumn() || 1;
    const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headers = rawHeaders.map(h => String(h).trim().toLowerCase());
    const dataValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const dateCol = headers.indexOf(DATE_COLUMN_NAME.toLowerCase());
    const companyCol = headers.indexOf(COMPANY_COLUMN_NAME.toLowerCase());
    const dbaCol = headers.indexOf(DBA_COLUMN_NAME.toLowerCase()); 
    const facilityTypeCol = headers.indexOf(FACILITY_TYPE_COLUMN_NAME.toLowerCase());
    const addressCol = headers.indexOf(ADDRESS_COLUMN_NAME.toLowerCase());
    const emailCol = headers.indexOf(EMAIL_COLUMN_NAME.toLowerCase());
    const phoneCol = headers.indexOf(PHONE_COLUMN_NAME.toLowerCase());
    const statusCol = headers.indexOf(STATUS_COLUMN_NAME.toLowerCase());

    const corporateCol = headers.indexOf(CORPORATE_COLUMN_NAME.toLowerCase());
    const leadGeneratorCol = headers.indexOf(LEAD_GENERATOR_COLUMN_NAME.toLowerCase());
    const accountOwnerCol = headers.indexOf(ACCOUNT_OWNER_COLUMN_NAME.toLowerCase()); 
    const adminNameCol = headers.indexOf(ADMIN_NAME_COLUMN_NAME.toLowerCase());
    const contactNameCol = headers.indexOf(CONTACT_NAME_COLUMN_NAME.toLowerCase());
    const contactTitleCol = headers.indexOf(CONTACT_TITLE_COLUMN_NAME.toLowerCase());
    const contactEmailCol = headers.indexOf(CONTACT_EMAIL_COLUMN_NAME.toLowerCase());
    const contactPhoneCol = headers.indexOf(CONTACT_PHONE_COLUMN_NAME.toLowerCase());
    const centralSupplyCol = headers.indexOf(CENTRAL_SUPPLY_COLUMN_NAME.toLowerCase());
    const patientSupplyCol = headers.indexOf(PATIENT_SUPPLY_COLUMN_NAME.toLowerCase());
    const partBBillingCol = headers.indexOf(PART_B_BILLING_COLUMN_NAME.toLowerCase());
    const hmoBillingCol = headers.indexOf(HMO_BILLING_COLUMN_NAME.toLowerCase());
    const cityCol = headers.indexOf(CITY_COLUMN_NAME.toLowerCase());
    const stateCol = headers.indexOf(STATE_COLUMN_NAME.toLowerCase());
    const zipCol = headers.indexOf(ZIP_COLUMN_NAME.toLowerCase());

    const formattedCompanies = dataValues.map((rowArr, index) => {
      const rowNum = index + 2; 
      let dateVal = (dateCol >= 0 && dateCol < rowArr.length) ? rowArr[dateCol] : "";

      if (dateVal instanceof Date) {
        try {
          dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch(e) {
          dateVal = "";
        }
      } else if (dateVal) {
        dateVal = String(dateVal);
      }

      let currentStatus = (statusCol >= 0 && statusCol < rowArr.length) ? String(rowArr[statusCol]).trim() : "";
      if(!currentStatus) currentStatus = DEFAULT_STATUS;

      return {
        row: rowNum,
        date: dateVal,
        callDate: dateVal,
        company: (companyCol >= 0 && companyCol < rowArr.length) ? String(rowArr[companyCol]).trim() : "",
        dba: (dbaCol >= 0 && dbaCol < rowArr.length) ? String(rowArr[dbaCol]).trim() : "", 
        facilityType: (facilityTypeCol >= 0 && facilityTypeCol < rowArr.length) ? String(rowArr[facilityTypeCol]).trim() : "",
        address: (addressCol >= 0 && addressCol < rowArr.length) ? String(rowArr[addressCol]).trim() : "",
        email: (emailCol >= 0 && emailCol < rowArr.length) ? String(rowArr[emailCol]).trim() : "",
        phone: (phoneCol >= 0 && phoneCol < rowArr.length) ? String(rowArr[phoneCol]).trim() : "",
        currentStatus: currentStatus,

        corporate: (corporateCol >= 0 && corporateCol < rowArr.length) ? String(rowArr[corporateCol]).trim() : "",
        leadGenerator: (leadGeneratorCol >= 0 && leadGeneratorCol < rowArr.length) ? String(rowArr[leadGeneratorCol]).trim() : "",
        accountOwner: (accountOwnerCol >= 0 && accountOwnerCol < rowArr.length) ? String(rowArr[accountOwnerCol]).trim() : "", 
        adminName: (adminNameCol >= 0 && adminNameCol < rowArr.length) ? String(rowArr[adminNameCol]).trim() : "",
        contactName: (contactNameCol >= 0 && contactNameCol < rowArr.length) ? String(rowArr[contactNameCol]).trim() : "",
        contactTitle: (contactTitleCol >= 0 && contactTitleCol < rowArr.length) ? String(rowArr[contactTitleCol]).trim() : "",
        contactEmail: (contactEmailCol >= 0 && contactEmailCol < rowArr.length) ? String(rowArr[contactEmailCol]).trim() : "",
        contactPhone: (contactPhoneCol >= 0 && contactPhoneCol < rowArr.length) ? String(rowArr[contactPhoneCol]).trim() : "",
        centralSupply: (centralSupplyCol >= 0 && centralSupplyCol < rowArr.length) ? !!rowArr[centralSupplyCol] : false,
        patientSupply: (patientSupplyCol >= 0 && patientSupplyCol < rowArr.length) ? !!rowArr[patientSupplyCol] : false,
        partBBilling: (partBBillingCol >= 0 && partBBillingCol < rowArr.length) ? !!rowArr[partBBillingCol] : false,
        hmoBilling: (hmoBillingCol >= 0 && hmoBillingCol < rowArr.length) ? !!rowArr[hmoBillingCol] : false,
        city: (cityCol >= 0 && cityCol < rowArr.length) ? String(rowArr[cityCol]).trim() : "",
        state: (stateCol >= 0 && stateCol < rowArr.length) ? String(rowArr[stateCol]).trim() : "",
        zip: (zipCol >= 0 && zipCol < rowArr.length) ? String(rowArr[zipCol]).trim() : ""
      };
    });

    return {
      companies: formattedCompanies,
      statuses: STATUSES,
      outcomes: CALL_OUTCOMES
    };
  } catch (error) {
    throw new Error("Backend synchronization error: " + error.toString());
  }
}

function getTabData(tabKey) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheetName = "";
    if (tabKey === "clientDirectory") targetSheetName = CLIENT_DIRECTORY_SHEET_NAME;
    else if (tabKey === "cpm2024MasterList") targetSheetName = CPM_MASTER_LIST_SHEET_NAME;
    else if (tabKey === "homeHealth") targetSheetName = HOME_HEALTH_SHEET_NAME;
    else if (tabKey === "hospice") targetSheetName = HOSPICE_SHEET_NAME;
    else throw new Error("Invalid tab specified.");

    let sheet = ss.getSheetByName(targetSheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { headers: [], records: [], statuses: STATUSES, outcomes: CALL_OUTCOMES };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const lowerHeaders = headers.map(h => h.toLowerCase());

    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const recordData = {};

      headers.forEach((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        } else {
          val = val !== null && val !== undefined ? String(val).trim() : "";
        }
        recordData[h] = val;
      });

      let statusVal = "";
      let outcomeVal = "";

      const statusIdx = lowerHeaders.indexOf("status");
      const outcomeIdx = lowerHeaders.indexOf("outcome");

      if (tabKey === "clientDirectory") {
        const colIdx = statusIdx >= 0 ? statusIdx : outcomeIdx;
        statusVal = colIdx >= 0 ? String(row[colIdx] || DEFAULT_STATUS).trim() : DEFAULT_STATUS;
      } else {
        const colIdx = outcomeIdx >= 0 ? outcomeIdx : statusIdx;
        outcomeVal = colIdx >= 0 ? String(row[colIdx] || "To Call").trim() : "To Call";
      }

      const dateLoggedVal = recordData["Date Logged"] || recordData["date logged"] || recordData["Date"] || "";
      const notesVal = recordData["Notes"] || recordData["notes"] || recordData["Recent Call Notes"] || recordData["Call History Log"] || recordData["History Log"] || "";

      records.push({
        row: i + 1,
        dateLogged: dateLoggedVal,
        accountName: recordData["Account Name"] || recordData["account name"] || recordData["Company Name"] || recordData["company name"] || recordData["Company"] || "",
        tier: recordData["Tier"] || recordData["tier"] || "",
        address: recordData["Address"] || recordData["address"] || "",
        city: recordData["City"] || recordData["city"] || "",
        state: recordData["State"] || recordData["state"] || "",
        county: recordData["County"] || recordData["county"] || "",
        zip: recordData["Zip"] || recordData["zip"] || recordData["Zip Code"] || "",
        phone: recordData["Phone Number"] || recordData["phone number"] || recordData["Phone"] || "",
        email: recordData["Email Address"] || recordData["email address"] || recordData["Email"] || recordData["email"] || "",
        status: statusVal,
        outcome: outcomeVal,
        notes: notesVal,
        _dataMap: recordData
      });
    }

    return { headers: headers, records: records, statuses: STATUSES, outcomes: CALL_OUTCOMES };
  } catch (err) {
    return { headers: [], records: [], statuses: STATUSES, outcomes: CALL_OUTCOMES, error: err.toString() };
  }
}

function saveTabCallRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheetName = "";

    if (data.tab === "clientDirectory") targetSheetName = CLIENT_DIRECTORY_SHEET_NAME;
    else if (data.tab === "cpm2024MasterList") targetSheetName = CPM_MASTER_LIST_SHEET_NAME;
    else if (data.tab === "homeHealth") targetSheetName = HOME_HEALTH_SHEET_NAME;
    else if (data.tab === "hospice") targetSheetName = HOSPICE_SHEET_NAME;
    else throw new Error("Invalid tab specified.");

    let sheet = ss.getSheetByName(targetSheetName);
    const statusColName = (data.tab === "clientDirectory") ? "Status" : "Outcome";

    if (!sheet) {
      sheet = ss.insertSheet(targetSheetName);
      sheet.appendRow(["Date Logged", "Account Name", "Tier", "Address", "City", "State", "County", "Zip", "Phone Number", "Email Address", statusColName, "Notes"]);
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold").setBackground("#F1F5F9");
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");
    const callDateFormatted = data.callDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const currentUser = data.user || getActiveCRMUser();
    const formattedNotes = data.notes ? `[Call Date: ${callDateFormatted} | ${timestamp} - ${currentUser}]: ${data.notes.trim()}` : "";

    const statusOrOutcomeVal = (data.tab === "clientDirectory") 
      ? (data.status || DEFAULT_STATUS) 
      : (data.outcome || "To Call");

    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());

    if (rawHeaders.length > 0) {
      const rowArr = new Array(rawHeaders.length).fill("");
      const setColVal = (headerName, val) => {
        const idx = rawHeaders.indexOf(headerName.toLowerCase());
        if (idx >= 0) rowArr[idx] = val;
      };

      setColVal("date logged", callDateFormatted);
      setColVal("date", callDateFormatted);
      setColVal("account name", data.accountName || "");
      setColVal("company name", data.accountName || "");
      setColVal("tier", data.tier || "");
      setColVal("address", data.address || "");
      setColVal("city", data.city || "");
      setColVal("state", data.state || "");
      setColVal("county", data.county || "");
      setColVal("zip", data.zip || "");
      setColVal("phone number", data.phone || "");
      setColVal("email address", data.email || "");
      setColVal(statusColName.toLowerCase(), statusOrOutcomeVal);

      // Dynamically match any note/history column headers
      rawHeaders.forEach((h, idx) => {
        if (h.includes("note") || h.includes("history")) {
          rowArr[idx] = formattedNotes;
        }
      });

      sheet.appendRow(rowArr);
    } else {
      sheet.appendRow([
        callDateFormatted,
        data.accountName || "",
        data.tier || "",
        data.address || "",
        data.city || "",
        data.state || "",
        data.county || "",
        data.zip || "",
        data.phone || "",
        data.email || "",
        statusOrOutcomeVal,
        formattedNotes
      ]);
    }

    return { success: true, message: "Call logged successfully!" };
  } catch (err) {
    return { success: false, message: "Failed to log call: " + err.toString() };
  }
}

function updateTabCallRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheetName = "";

    if (data.tab === "clientDirectory") targetSheetName = CLIENT_DIRECTORY_SHEET_NAME;
    else if (data.tab === "cpm2024MasterList") targetSheetName = CPM_MASTER_LIST_SHEET_NAME;
    else if (data.tab === "homeHealth") targetSheetName = HOME_HEALTH_SHEET_NAME;
    else if (data.tab === "hospice") targetSheetName = HOSPICE_SHEET_NAME;
    else throw new Error("Invalid tab specified.");

    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error("Sheet not found.");

    const row = parseInt(data.row, 10);
    if (!row || isNaN(row) || row <= 1) throw new Error("Invalid row selected.");

    const statusColName = (data.tab === "clientDirectory") ? "Status" : "Outcome";
    const statusOrOutcomeVal = (data.tab === "clientDirectory") 
      ? (data.status || DEFAULT_STATUS) 
      : (data.outcome || "To Call");

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");
    const callDateFormatted = data.callDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const currentUser = data.user || getActiveCRMUser();

    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());

    const updateCol = (headerName, val) => {
      const idx = rawHeaders.indexOf(headerName.toLowerCase());
      if (idx >= 0) {
        sheet.getRange(row, idx + 1).setValue(val);
      }
    };

    if (data.callDate) {
      updateCol("date logged", callDateFormatted);
      updateCol("date", callDateFormatted);
    }

    updateCol("account name", data.accountName || "");
    updateCol("company name", data.accountName || "");
    updateCol("tier", data.tier || "");
    updateCol("address", data.address || "");
    updateCol("city", data.city || "");
    updateCol("state", data.state || "");
    updateCol("county", data.county || "");
    updateCol("zip", data.zip || "");
    updateCol("phone number", data.phone || "");
    updateCol("phone", data.phone || "");
    updateCol("email address", data.email || "");
    updateCol("email", data.email || "");
    updateCol(statusColName.toLowerCase(), statusOrOutcomeVal);

    if (data.notes && data.notes.trim() !== "") {
      const formattedNote = `[Call Date: ${callDateFormatted} | ${timestamp} - ${currentUser}]: ${data.notes.trim()}`;
      
      let noteUpdated = false;
      rawHeaders.forEach((h, idx) => {
        if (h.includes("note") || h.includes("history")) {
          const currentNotesVal = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
          const newNotesVal = currentNotesVal ? `${formattedNote}\n${currentNotesVal}` : formattedNote;
          sheet.getRange(row, idx + 1).setValue(newNotesVal);
          noteUpdated = true;
        }
      });

      if (!noteUpdated) {
        // Fallback to appended column if no notes column header exists
        const lastCol = sheet.getLastColumn();
        sheet.getRange(1, lastCol + 1).setValue("Notes").setFontWeight("bold");
        sheet.getRange(row, lastCol + 1).setValue(formattedNote);
      }
    }

    return { success: true, message: "Record updated successfully!" };
  } catch (err) {
    return { success: false, message: "Failed to update record: " + err.toString() };
  }
}

function getRowDataByNumber(row) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: "CRM sheet not found.", statuses: STATUSES };

    const lastCol = sheet.getLastColumn() || 1;
    const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headers = rawHeaders.map(h => String(h).trim().toLowerCase());

    const dateCol = headers.indexOf(DATE_COLUMN_NAME.toLowerCase()) + 1;
    const companyCol = headers.indexOf(COMPANY_COLUMN_NAME.toLowerCase()) + 1;
    const dbaCol = headers.indexOf(DBA_COLUMN_NAME.toLowerCase()) + 1; 
    const facilityTypeCol = headers.indexOf(FACILITY_TYPE_COLUMN_NAME.toLowerCase()) + 1;
    const addressCol = headers.indexOf(ADDRESS_COLUMN_NAME.toLowerCase()) + 1;
    const emailCol = headers.indexOf(EMAIL_COLUMN_NAME.toLowerCase()) + 1;
    const phoneCol = headers.indexOf(PHONE_COLUMN_NAME.toLowerCase()) + 1;
    const statusCol = headers.indexOf(STATUS_COLUMN_NAME.toLowerCase()) + 1;
    const recentNotesCol = headers.indexOf(RECENT_NOTES_COLUMN_NAME.toLowerCase()) + 1;
    const historyCol = headers.indexOf(HISTORY_COLUMN_NAME.toLowerCase()) + 1;

    const corporateCol = headers.indexOf(CORPORATE_COLUMN_NAME.toLowerCase()) + 1;
    const leadGeneratorCol = headers.indexOf(LEAD_GENERATOR_COLUMN_NAME.toLowerCase()) + 1;
    const accountOwnerCol = headers.indexOf(ACCOUNT_OWNER_COLUMN_NAME.toLowerCase()) + 1; 
    const adminNameCol = headers.indexOf(ADMIN_NAME_COLUMN_NAME.toLowerCase()) + 1;
    const contactNameCol = headers.indexOf(CONTACT_NAME_COLUMN_NAME.toLowerCase()) + 1;
    const contactTitleCol = headers.indexOf(CONTACT_TITLE_COLUMN_NAME.toLowerCase()) + 1;
    const contactEmailCol = headers.indexOf(CONTACT_EMAIL_COLUMN_NAME.toLowerCase()) + 1;
    const centralSupplyCol = headers.indexOf(CENTRAL_SUPPLY_COLUMN_NAME.toLowerCase()) + 1;
    const patientSupplyCol = headers.indexOf(PATIENT_SUPPLY_COLUMN_NAME.toLowerCase()) + 1;
    const partBBillingCol = headers.indexOf(PART_B_BILLING_COLUMN_NAME.toLowerCase()) + 1;
    const hmoBillingCol = headers.indexOf(HMO_BILLING_COLUMN_NAME.toLowerCase()) + 1;
    const cityCol = headers.indexOf(CITY_COLUMN_NAME.toLowerCase()) + 1;
    const stateCol = headers.indexOf(STATE_COLUMN_NAME.toLowerCase()) + 1;
    const zipCol = headers.indexOf(ZIP_COLUMN_NAME.toLowerCase()) + 1;

    let dateVal = "";
    if (dateCol > 0) {
      const tempDate = sheet.getRange(row, dateCol).getValue();
      if (tempDate instanceof Date) {
        dateVal = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (tempDate) {
        dateVal = String(tempDate);
      }
    }

    let currentStatus = statusCol > 0 ? String(sheet.getRange(row, statusCol).getValue()).trim() : "";
    if(!currentStatus) currentStatus = DEFAULT_STATUS;

    return {
      row: row,
      leadName: (companyCol > 0 ? String(sheet.getRange(row, companyCol).getValue()).trim() : "") || `Row ${row}`,
      date: dateVal,
      callDate: dateVal,
      company: companyCol > 0 ? String(sheet.getRange(row, companyCol).getValue()).trim() : "",
      dba: dbaCol > 0 ? String(sheet.getRange(row, dbaCol).getValue()).trim() : "", 
      facilityType: facilityTypeCol > 0 ? String(sheet.getRange(row, facilityTypeCol).getValue()).trim() : "",
      address: addressCol > 0 ? String(sheet.getRange(row, addressCol).getValue()).trim() : "",
      email: emailCol > 0 ? String(sheet.getRange(row, emailCol).getValue()).trim() : "",
      phone: phoneCol > 0 ? String(sheet.getRange(row, phoneCol).getValue()).trim() : "",
      currentStatus: currentStatus,
      recentNotes: recentNotesCol > 0 ? String(sheet.getRange(row, recentNotesCol).getValue()).trim() : "",
      callHistory: historyCol > 0 ? String(sheet.getRange(row, historyCol).getValue()).trim() : "",
      statuses: STATUSES,

      corporate: corporateCol > 0 ? String(sheet.getRange(row, corporateCol).getValue()).trim() : "",
      leadGenerator: leadGeneratorCol > 0 ? String(sheet.getRange(row, leadGeneratorCol).getValue()).trim() : "",
      accountOwner: accountOwnerCol > 0 ? String(sheet.getRange(row, accountOwnerCol).getValue()).trim() : "", 
      adminName: adminNameCol > 0 ? String(sheet.getRange(row, adminNameCol).getValue()).trim() : "",
      contactName: contactNameCol > 0 ? String(sheet.getRange(row, contactNameCol).getValue()).trim() : "",
      contactTitle: contactTitleCol > 0 ? String(sheet.getRange(row, contactTitleCol).getValue()).trim() : "",
      contactEmail: contactEmailCol > 0 ? String(sheet.getRange(row, contactEmailCol).getValue()).trim() : "",
      centralSupply: centralSupplyCol > 0 ? !!sheet.getRange(row, centralSupplyCol).getValue() : false,
      patientSupply: patientSupplyCol > 0 ? !!sheet.getRange(row, patientSupplyCol).getValue() : false,
      partBBilling: partBBillingCol > 0 ? !!sheet.getRange(row, partBBillingCol).getValue() : false,
      hmoBilling: hmoBillingCol > 0 ? !!sheet.getRange(row, hmoBillingCol).getValue() : false,
      city: cityCol > 0 ? String(sheet.getRange(row, cityCol).getValue()).trim() : "",
      state: stateCol > 0 ? String(sheet.getRange(row, stateCol).getValue()).trim() : "",
      zip: zipCol > 0 ? String(sheet.getRange(row, zipCol).getValue()).trim() : ""
    };
  } catch (error) {
    throw new Error("Failed to load workspace data: " + error.toString());
  }
}

function addNewLeadFromSidebar(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setupCRM();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = rawHeaders.map(h => String(h).trim().toLowerCase());
  const newRowData = new Array(headers.length).fill("");

  const dateCol = headers.indexOf(DATE_COLUMN_NAME.toLowerCase());
  const companyCol = headers.indexOf(COMPANY_COLUMN_NAME.toLowerCase());
  const dbaCol = headers.indexOf(DBA_COLUMN_NAME.toLowerCase()); 
  const facilityTypeCol = headers.indexOf(FACILITY_TYPE_COLUMN_NAME.toLowerCase());
  const addressCol = headers.indexOf(ADDRESS_COLUMN_NAME.toLowerCase());
  const emailCol = headers.indexOf(EMAIL_COLUMN_NAME.toLowerCase());
  const phoneCol = headers.indexOf(PHONE_COLUMN_NAME.toLowerCase());
  const statusCol = headers.indexOf(STATUS_COLUMN_NAME.toLowerCase());
  const recentNotesCol = headers.indexOf(RECENT_NOTES_COLUMN_NAME.toLowerCase());
  const historyCol = headers.indexOf(HISTORY_COLUMN_NAME.toLowerCase());
  const callCountCol = headers.indexOf(CALL_COUNT_COLUMN_NAME.toLowerCase());
  const lastCalledCol = headers.indexOf(LAST_CALLED_COLUMN_NAME.toLowerCase());

  const corporateCol = headers.indexOf(CORPORATE_COLUMN_NAME.toLowerCase());
  const leadGeneratorCol = headers.indexOf(LEAD_GENERATOR_COLUMN_NAME.toLowerCase());
  const accountOwnerCol = headers.indexOf(ACCOUNT_OWNER_COLUMN_NAME.toLowerCase()); 
  const adminNameCol = headers.indexOf(ADMIN_NAME_COLUMN_NAME.toLowerCase());
  const contactNameCol = headers.indexOf(CONTACT_NAME_COLUMN_NAME.toLowerCase());
  const contactTitleCol = headers.indexOf(CONTACT_TITLE_COLUMN_NAME.toLowerCase());
  const contactEmailCol = headers.indexOf(CONTACT_EMAIL_COLUMN_NAME.toLowerCase());
  const contactPhoneCol = headers.indexOf(CONTACT_PHONE_COLUMN_NAME.toLowerCase());
  const centralSupplyCol = headers.indexOf(CENTRAL_SUPPLY_COLUMN_NAME.toLowerCase());
  const patientSupplyCol = headers.indexOf(PATIENT_SUPPLY_COLUMN_NAME.toLowerCase());
  const partBBillingCol = headers.indexOf(PART_B_BILLING_COLUMN_NAME.toLowerCase());
  const hmoBillingCol = headers.indexOf(HMO_BILLING_COLUMN_NAME.toLowerCase());
  const cityCol = headers.indexOf(CITY_COLUMN_NAME.toLowerCase());
  const stateCol = headers.indexOf(STATE_COLUMN_NAME.toLowerCase());
  const zipCol = headers.indexOf(ZIP_COLUMN_NAME.toLowerCase());

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");
  const callDateFormatted = data.callDate || data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const currentUser = data.user || getActiveCRMUser();

  if (dateCol >= 0 && callDateFormatted) {
    const parts = callDateFormatted.split('-');
    if(parts.length === 3) {
      newRowData[dateCol] = new Date(parts[0], parts[1] - 1, parts[2]);
    }
  } else if (dateCol >= 0) {
    newRowData[dateCol] = new Date();
  }

  if (companyCol >= 0) newRowData[companyCol] = data.company || "";
  if (dbaCol >= 0) newRowData[dbaCol] = data.dba || ""; 
  if (facilityTypeCol >= 0) newRowData[facilityTypeCol] = data.facilityType || "";
  if (addressCol >= 0) newRowData[addressCol] = data.address || "";
  if (emailCol >= 0) newRowData[emailCol] = data.email || "";
  if (phoneCol >= 0) newRowData[phoneCol] = data.phone || "";
  if (statusCol >= 0) newRowData[statusCol] = data.status || DEFAULT_STATUS;

  if (corporateCol >= 0) newRowData[corporateCol] = data.corporate || "";
  if (leadGeneratorCol >= 0) newRowData[leadGeneratorCol] = data.leadGenerator || "";
  if (accountOwnerCol >= 0) newRowData[accountOwnerCol] = data.accountOwner || ""; 
  if (adminNameCol >= 0) newRowData[adminNameCol] = data.adminName || "";
  if (contactNameCol >= 0) newRowData[contactNameCol] = data.contactName || "";
  if (contactTitleCol >= 0) newRowData[contactTitleCol] = data.contactTitle || "";
  if (contactEmailCol >= 0) newRowData[contactEmailCol] = data.contactEmail || "";
  if (contactPhoneCol >= 0) newRowData[contactPhoneCol] = data.contactPhone || "";
  if (centralSupplyCol >= 0) newRowData[centralSupplyCol] = data.centralSupply || false;
  if (patientSupplyCol >= 0) newRowData[patientSupplyCol] = data.patientSupply || false;
  if (partBBillingCol >= 0) newRowData[partBBillingCol] = data.partBBilling || false;
  if (hmoBillingCol >= 0) newRowData[hmoBillingCol] = data.hmoBilling || false;
  if (cityCol >= 0) newRowData[cityCol] = data.city || "";
  if (stateCol >= 0) newRowData[stateCol] = data.state || "";
  if (zipCol >= 0) newRowData[zipCol] = data.zip || "";

  if (data.note && data.note.trim() !== "") {
    const formattedNote = `[Call Date: ${callDateFormatted} | ${timestamp} - ${currentUser}]: ${data.note.trim()}`;
    if (recentNotesCol >= 0) newRowData[recentNotesCol] = formattedNote;
    if (historyCol >= 0) newRowData[historyCol] = formattedNote;
    if (callCountCol >= 0) newRowData[callCountCol] = 1;
    if (lastCalledCol >= 0) newRowData[lastCalledCol] = new Date();
  } else {
    if (callCountCol >= 0) newRowData[callCountCol] = 0;
  }

  sheet.appendRow(newRowData);
  const newRowNum = sheet.getLastRow();

  if (statusCol >= 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
    sheet.getRange(newRowNum, statusCol + 1).setDataValidation(rule);
    colorCodeRow(sheet, newRowNum, data.status || DEFAULT_STATUS);
  }

  return { message: "Company added successfully!", rowNumber: newRowNum };
}

function updateLeadFromSidebar(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const row = parseInt(data.row);
  if (!row || isNaN(row)) return "Error: No valid row selected to update.";

  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = rawHeaders.map(h => String(h).trim().toLowerCase());

  const dateCol = headers.indexOf(DATE_COLUMN_NAME.toLowerCase()) + 1;
  const companyCol = headers.indexOf(COMPANY_COLUMN_NAME.toLowerCase()) + 1;
  const dbaCol = headers.indexOf(DBA_COLUMN_NAME.toLowerCase()) + 1; 
  const facilityTypeCol = headers.indexOf(FACILITY_TYPE_COLUMN_NAME.toLowerCase()) + 1;
  const addressCol = headers.indexOf(ADDRESS_COLUMN_NAME.toLowerCase()) + 1;
  const emailCol = headers.indexOf(EMAIL_COLUMN_NAME.toLowerCase()) + 1;
  const phoneCol = headers.indexOf(PHONE_COLUMN_NAME.toLowerCase()) + 1;
  const statusCol = headers.indexOf(STATUS_COLUMN_NAME.toLowerCase()) + 1;

  const lastCalledCol = headers.indexOf(LAST_CALLED_COLUMN_NAME.toLowerCase()) + 1;
  const callCountCol = headers.indexOf(CALL_COUNT_COLUMN_NAME.toLowerCase()) + 1;
  const recentNotesCol = headers.indexOf(RECENT_NOTES_COLUMN_NAME.toLowerCase()) + 1;
  const historyCol = headers.indexOf(HISTORY_COLUMN_NAME.toLowerCase()) + 1;

  const corporateCol = headers.indexOf(CORPORATE_COLUMN_NAME.toLowerCase()) + 1;
  const leadGeneratorCol = headers.indexOf(LEAD_GENERATOR_COLUMN_NAME.toLowerCase()) + 1;
  const accountOwnerCol = headers.indexOf(ACCOUNT_OWNER_COLUMN_NAME.toLowerCase()) + 1; 
  const adminNameCol = headers.indexOf(ADMIN_NAME_COLUMN_NAME.toLowerCase()) + 1;
  const contactNameCol = headers.indexOf(CONTACT_NAME_COLUMN_NAME.toLowerCase()) + 1;
  const contactTitleCol = headers.indexOf(CONTACT_TITLE_COLUMN_NAME.toLowerCase()) + 1;
  const contactEmailCol = headers.indexOf(CONTACT_EMAIL_COLUMN_NAME.toLowerCase()) + 1;
  const contactPhoneCol = headers.indexOf(CONTACT_PHONE_COLUMN_NAME.toLowerCase()) + 1;
  const centralSupplyCol = headers.indexOf(CENTRAL_SUPPLY_COLUMN_NAME.toLowerCase()) + 1;
  const patientSupplyCol = headers.indexOf(PATIENT_SUPPLY_COLUMN_NAME.toLowerCase()) + 1;
  const partBBillingCol = headers.indexOf(PART_B_BILLING_COLUMN_NAME.toLowerCase()) + 1;
  const hmoBillingCol = headers.indexOf(HMO_BILLING_COLUMN_NAME.toLowerCase()) + 1;
  const cityCol = headers.indexOf(CITY_COLUMN_NAME.toLowerCase()) + 1;
  const stateCol = headers.indexOf(STATE_COLUMN_NAME.toLowerCase()) + 1;
  const zipCol = headers.indexOf(ZIP_COLUMN_NAME.toLowerCase()) + 1;

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");
  const callDateFormatted = data.callDate || data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const currentUser = data.user || getActiveCRMUser();

  if (dateCol > 0 && callDateFormatted) {
    const parts = callDateFormatted.split('-');
    if(parts.length === 3) {
      sheet.getRange(row, dateCol).setValue(new Date(parts[0], parts[1] - 1, parts[2]));
    }
  }
  if (companyCol > 0) sheet.getRange(row, companyCol).setValue(data.company || "");
  if (dbaCol > 0) sheet.getRange(row, dbaCol).setValue(data.dba || ""); 
  if (facilityTypeCol > 0) sheet.getRange(row, facilityTypeCol).setValue(data.facilityType || "");
  if (addressCol > 0) sheet.getRange(row, addressCol).setValue(data.address || "");
  if (emailCol > 0) sheet.getRange(row, emailCol).setValue(data.email || "");
  if (phoneCol > 0) sheet.getRange(row, phoneCol).setValue(data.phone || "");

  if (corporateCol > 0) sheet.getRange(row, corporateCol).setValue(data.corporate || "");
  if (leadGeneratorCol > 0) sheet.getRange(row, leadGeneratorCol).setValue(data.leadGenerator || "");
  if (accountOwnerCol > 0) sheet.getRange(row, accountOwnerCol).setValue(data.accountOwner || ""); 
  if (adminNameCol > 0) sheet.getRange(row, adminNameCol).setValue(data.adminName || "");
  if (contactNameCol > 0) sheet.getRange(row, contactNameCol).setValue(data.contactName || "");
  if (contactTitleCol > 0) sheet.getRange(row, contactTitleCol).setValue(data.contactTitle || ""); 
  if (contactEmailCol > 0) sheet.getRange(row, contactEmailCol).setValue(data.contactEmail || "");
  if (contactPhoneCol > 0) sheet.getRange(row, contactPhoneCol).setValue(data.contactPhone || "");
  if (centralSupplyCol > 0) sheet.getRange(row, centralSupplyCol).setValue(data.centralSupply || false);
  if (patientSupplyCol > 0) sheet.getRange(row, patientSupplyCol).setValue(data.patientSupply || false);
  if (partBBillingCol > 0) sheet.getRange(row, partBBillingCol).setValue(data.partBBilling || false);
  if (hmoBillingCol > 0) sheet.getRange(row, hmoBillingCol).setValue(data.hmoBilling || false);
  if (cityCol > 0) sheet.getRange(row, cityCol).setValue(data.city || "");
  if (stateCol > 0) sheet.getRange(row, stateCol).setValue(data.state || "");
  if (zipCol > 0) sheet.getRange(row, zipCol).setValue(data.zip || "");

  if (statusCol > 0 && data.status) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
    sheet.getRange(row, statusCol).setDataValidation(rule);
    sheet.getRange(row, statusCol).setValue(data.status);
    colorCodeRow(sheet, row, data.status);
  }

  if (data.note && data.note.trim() !== "") {
    const formattedNote = `[Call Date: ${callDateFormatted} | ${timestamp} - ${currentUser}]: ${data.note.trim()}`;
    if (lastCalledCol > 0) sheet.getRange(row, lastCalledCol).setValue(new Date());
    if (callCountCol > 0) {
      const currentCountCell = sheet.getRange(row, callCountCol);
      currentCountCell.setValue((Number(currentCountCell.getValue()) || 0) + 1);
    }
    if (recentNotesCol > 0) sheet.getRange(row, recentNotesCol).setValue(formattedNote);
    if (historyCol > 0) {
      const historyCell = sheet.getRange(row, historyCol);
      const oldHistory = historyCell.getValue();
      historyCell.setValue(oldHistory ? `${formattedNote}\n${oldHistory}` : formattedNote);
    }
  }
  return "Lead updated successfully!";
}

function deleteLeadFromSidebar(row) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const rowNum = parseInt(row, 10);
    if (!sheet) throw new Error("CRM Spreadsheet tab not found.");
    if (isNaN(rowNum) || rowNum <= 1 || rowNum > sheet.getLastRow()) {
      throw new Error("Invalid record line selected.");
    }

    sheet.deleteRow(rowNum);
    return "Record deleted successfully!";
  } catch (error) {
    throw new Error("Failed to delete record: " + error.toString());
  }
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const statusCol = headers.indexOf(STATUS_COLUMN_NAME.toLowerCase()) + 1;

  if (range.getColumn() === statusCol && range.getRow() > 1) {
    colorCodeRow(sheet, range.getRow(), range.getValue());
  }
}

function colorCodeRow(sheet, row, status) {
  const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  let backgroundColor = "#FFFFFF";
  let textColor = "#000000";
  let fontWeight = "normal";

  switch(status) {
    case "Potential Lead / Attempting Contact":
    case "Engaged / Connected":
    case "Negotiation / In Review":
    case "Paused":
      backgroundColor = "#EAF1F7"; 
      break;
    case "Closed Won":
    case "Active Client":
    case "Onboarding in Progress":
      backgroundColor = "#C6DCED"; 
      break;
    case "Disqualified / Unfit":
    case "Closed Lost":
    case "At Risk":
    case "Past Client":
      backgroundColor = "#A2C7E3"; 
      break;
  }

  range.setBackground(backgroundColor);
  range.setFontColor(textColor).setFontWeight(fontWeight);
}

function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if(sheet.getLastColumn() === 0) {
    sheet.appendRow([
      DATE_COLUMN_NAME, 
      CORPORATE_COLUMN_NAME,
      LEAD_GENERATOR_COLUMN_NAME,
      ACCOUNT_OWNER_COLUMN_NAME, 
      ADMIN_NAME_COLUMN_NAME,
      COMPANY_COLUMN_NAME, 
      DBA_COLUMN_NAME, 
      FACILITY_TYPE_COLUMN_NAME,
      CONTACT_NAME_COLUMN_NAME,
      CONTACT_TITLE_COLUMN_NAME,
      CONTACT_EMAIL_COLUMN_NAME,
      CONTACT_PHONE_COLUMN_NAME,
      CENTRAL_SUPPLY_COLUMN_NAME,
      PATIENT_SUPPLY_COLUMN_NAME,
      PART_B_BILLING_COLUMN_NAME,
      HMO_BILLING_COLUMN_NAME,
      ADDRESS_COLUMN_NAME, 
      CITY_COLUMN_NAME,
      STATE_COLUMN_NAME,
      ZIP_COLUMN_NAME,
      EMAIL_COLUMN_NAME, 
      PHONE_COLUMN_NAME, 
      STATUS_COLUMN_NAME, 
      LAST_CALLED_COLUMN_NAME, 
      CALL_COUNT_COLUMN_NAME, 
      RECENT_NOTES_COLUMN_NAME, 
      HISTORY_COLUMN_NAME
    ]);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  let statusCol = headers.indexOf(STATUS_COLUMN_NAME.toLowerCase()) + 1;

  if (statusCol > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
    sheet.getRange(2, statusCol, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  }
}