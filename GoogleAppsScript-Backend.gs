/**
 * MajengoSmart Pro - Google Apps Script Backend
 * 
 * This script acts as the backend for the MajengoSmart Pro system.
 * It handles:
 * - Email sending (invoices, reminders, notifications)
 * - Google Drive storage and archiving
 * - Data backup and restoration
 * 
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Go to script.google.com
 * 2. Create a new project
 * 3. Copy this entire code into Code.gs
 * 4. Click Deploy → New Deployment
 * 5. Select "Web app" as deployment type
 * 6. Set "Execute as: Me"
 * 7. Set "Who has access: Anyone" (for your domain or anyone)
 * 8. Click "Deploy"
 * 9. Copy the Web App URL
 * 10. Paste it in your MajengoSmart Settings → Email Settings
 */

// ====== CONFIGURATION ======
const CONFIG = {
  COMPANY_NAME: 'MajengoSmart Property Management',
  FROM_EMAIL: 'billing@majengosmart.co.ke', // Change to your company email
  FROM_NAME: 'MajengoSmart Billing',
  DRIVE_FOLDER_NAME: 'MajengoSmart Data',
  DRIVE_ROOT_FOLDER_ID: '1KgJLD4ORCEFJc6ZLj8R6NhUiwU5-8K95', // TENANT MANAGEMENT SYSTEM folder
  TIMEZONE: 'Africa/Nairobi'
};

// ====== MAIN WEB APP HANDLER ======
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    // Route to appropriate handler
    switch(action) {
            case 'create_team':
              return createTeam(params);
            case 'list_teams':
              return listTeams(params);
            case 'create_user_account':
              return createUserAccount(params);
            case 'list_users':
              return listUsers(params);
            case 'activate_user':
              return setUserStatus(params, 'Active');
            case 'deactivate_user':
              return setUserStatus(params, 'Inactive');
            case 'activate_team':
              return setTeamStatus(params, 'Active');
            case 'deactivate_team':
              return setTeamStatus(params, 'Inactive');
            case 'set_team_permissions':
              return setTeamPermissions(params);
            case 'schedule_auto_send':
              return scheduleAutoSend(params);
            case 'generate_billing':
              return generateBilling(params);
            case 'run_billing_manual':
              return runBillingManual(params);
            case 'run_scheduled_sends':
              return runScheduledSends(params);
            case 'test_drive':
              return testDriveConnection(params);
            case 'init_coa':
              return initializeChartOfAccounts(params);
            case 'get_coa':
              return getChartOfAccounts(params);
            case 'generate_pl':
              return generateProfitLoss(params);
            case 'generate_bs':
              return generateBalanceSheet(params);
            case 'generate_tb':
              return generateTrialBalance(params);
            case 'get_monthly_charts':
              return getMonthlyCharts(params);

      case 'send_email':
        return sendEmail(params);
      case 'send_invoice':
        return sendInvoiceEmail(params);
      case 'create_folders':
        return createStorageFolders(params);
      case 'archive_tenant':
        return archiveTenant(params);
      case 'backup_data':
        return backupData(params);
      case 'init_sheets':
        return initSheets(params);
      case 'create_sheet_in_folder':
        return createSheetInFolder(params);
      case 'submit_repair_request':
        return submitRepairRequest(params);
      case 'review_repair_request':
        return reviewRepairRequest(params);
      case 'push_all':
        return pushAll(params);
      case 'push_customers':
        return pushCustomers(params);
      case 'push_invoices':
        return pushInvoices(params);
      case 'push_payments':
        return pushPayments(params);
      case 'push_accounts':
        return pushAccounts(params);
      case 'pull_all':
        return pullAll(params);
      case 'pull_customers':
        return pullCustomers(params);
      case 'pull_invoices':
        return pullInvoices(params);
      case 'pull_payments':
        return pullPayments(params);
      case 'pull_accounts':
        return pullAccounts(params);
      case 'import_monthly_invoices':
        return importMonthlyInvoices(params);
      case 'restore_data':
        return restoreData(params);
      case 'format_all_sheets':
        return formatAllSheets(params);
      case 'save_meta':
        return saveMetaData(params);
      default:
        return createResponse(false, 'Unknown action: ' + action);
    }
  } catch (error) {
    return createResponse(false, 'Error: ' + error.toString());
  }
}

/**
 * Initialize required sheets inside an existing spreadsheet (by ID or URL).
 * params: { spreadsheetId: string } - can be ID or full URL
 */
function initSheets(params) {
  try {
    let spreadsheetId = params.spreadsheetId;
    if(!spreadsheetId) return createResponse(false, 'spreadsheetId is required');

    // Extract ID from URL if full URL is provided
    if(spreadsheetId.includes('docs.google.com')) {
      const match = spreadsheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if(match && match[1]) {
        spreadsheetId = match[1];
      } else {
        return createResponse(false, 'Invalid Google Sheets URL');
      }
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    createDefaultSheets(ss);

    return createResponse(true, 'Sheets initialized', { spreadsheetId: ss.getId(), url: ss.getUrl() });
  } catch (error) {
    return createResponse(false, 'Failed to initialize sheets: ' + error.toString());
  }
}

/**
 * Save / update key-value pairs in the _meta sheet.
 * params: { spreadsheetId, data: { key: value, ... } }
 * e.g. { data: { companyName: 'Acme', companyEmail: 'info@acme.com' } }
 */
function saveMetaData(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    let metaSheet = ss.getSheetByName('_meta');
    if (!metaSheet) {
      metaSheet = ss.insertSheet('_meta');
    }
    const data = params.data || {};
    const keys = Object.keys(data);
    if (keys.length === 0) return createResponse(false, 'No data provided');

    // Read existing rows
    const existing = metaSheet.getDataRange().getValues(); // [[key, value], ...]
    const keyRowMap = {}; // key -> 1-based row index
    existing.forEach(function(row, i) {
      if (row[0]) keyRowMap[String(row[0]).trim()] = i + 1;
    });

    keys.forEach(function(k) {
      const v = data[k];
      if (keyRowMap[k]) {
        // Update existing row
        metaSheet.getRange(keyRowMap[k], 2).setValue(v);
      } else {
        // Append new row
        const nextRow = metaSheet.getLastRow() + 1;
        metaSheet.getRange(nextRow, 1).setValue(k);
        metaSheet.getRange(nextRow, 2).setValue(v);
        keyRowMap[k] = nextRow;
      }
    });

    return createResponse(true, 'Meta data saved', { keys: keys });
  } catch (error) {
    return createResponse(false, 'Failed to save meta data: ' + error.toString());
  }
}

/**
 * Create a new spreadsheet in a Drive folder and initialize it.
 * params: { folderId?: string, name?: string }
 */
function createSheetInFolder(params) {
  try {
    const name = params.name || 'MajengoSmart_Init_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
    let folder;
    if(params.folderId) {
      folder = DriveApp.getFolderById(params.folderId);
    } else {
      const main = getOrCreateFolder(CONFIG.DRIVE_FOLDER_NAME);
      folder = getOrCreateFolder('Spreadsheets', main);
    }

    const ss = SpreadsheetApp.create(name);
    // Move the sheet file into the target folder
    const file = DriveApp.getFileById(ss.getId());
    folder.addFile(file);
    // Remove from root to avoid duplicate listing
    DriveApp.getRootFolder().removeFile(file);

    createDefaultSheets(ss);

    return createResponse(true, 'Spreadsheet created and initialized', { spreadsheetId: ss.getId(), url: ss.getUrl(), folderId: folder.getId() });
  } catch (error) {
    return createResponse(false, 'Failed to create sheet in folder: ' + error.toString());
  }
}

/**
 * Create the default set of sheets and headers used by the system.
 */
function createDefaultSheets(ss) {
  // Helper to create a sheet if missing and set header row
  function createSheetIfNotExists(spreadsheet, name, headers) {
    let sheet = spreadsheet.getSheetByName(name);
    if(!sheet) sheet = spreadsheet.insertSheet(name);
    // Preserve existing data; only initialize an empty sheet.
    if(headers && headers.length && sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
    }
    return sheet;
  }

  // Define headers for each sheet
  const repairsHeaders = ['RequestID','TenantName','TenantEmail','Property','Unit','IssueDescription','Priority','DateLogged','Status','AssignedTo','ApprovedBy','ApprovedAt','Notes'];
  const tenantsHeaders = ['TenantID','FullName','IDNumber','Phone','Email','Property','Unit','MoveInDate','MonthlyRent','Status','EmergencyContact'];
  const rentHeaders = ['PaymentID','TenantName','Unit','Amount','Method','Reference','Date','Status','RecordedBy','Notes'];
  const invoicesHeaders = ['InvoiceID','TenantName','TenantEmail','Property','Unit','Amount','IssueDate','DueDate','Status','SentAt','PaymentRef','Notes'];
  const maintenanceApprovalHeaders = ['RequestID','TenantName','Unit','Issue','Priority','DateLogged','Status','ManagerAction','ActionBy','ActionDate','Notes'];
  const tenantStatementsHeaders = ['TenantID','TenantName','PeriodStart','PeriodEnd','OpeningBalance','Charges','Payments','ClosingBalance'];
  const agingTenantsHeaders = ['TenantID','TenantName','Current','30-59','60-89','90+','TotalDue'];
  const agingSuppliersHeaders = ['SupplierID','SupplierName','Current','30-59','60-89','90+','TotalDue'];

  const teamsHeaders = ['TeamID','TeamName','Permissions','CreatedAt','Notes','Status'];
  const usersHeaders = ['UserID','Email','FullName','Team','Role','Password','CreatedAt','Status'];

  createSheetIfNotExists(ss, 'Repairs', repairsHeaders);
  createSheetIfNotExists(ss, 'Tenants', tenantsHeaders);
  createSheetIfNotExists(ss, 'RentPayments', rentHeaders);
  createSheetIfNotExists(ss, 'Invoices', invoicesHeaders);
  createSheetIfNotExists(ss, 'MaintenanceApprovals', maintenanceApprovalHeaders);
  createSheetIfNotExists(ss, 'TenantStatements', tenantStatementsHeaders);
  createSheetIfNotExists(ss, 'Aging_Tenants', agingTenantsHeaders);
  createSheetIfNotExists(ss, 'Aging_Suppliers', agingSuppliersHeaders);

  createSheetIfNotExists(ss, 'Teams', teamsHeaders);
  const usersSheet = createSheetIfNotExists(ss, 'Users', usersHeaders);

  // If no users exist yet, create default Admin user
  try {
    const last = usersSheet.getLastRow();
    if (last <= 1) {
      const adminEmail = CONFIG.FROM_EMAIL || 'admin@example.com';
      usersSheet.appendRow(['U-ADMIN', adminEmail, 'Admin', 'Administrators', 'Admin', 'Admin@2026', new Date(), 'Active']);
    }
  } catch (e) {
    Logger.log('Failed to create default admin user: ' + e.toString());
  }

  ensureStatusColumn_(ss.getSheetByName('Teams'), 'Active');
  ensureStatusColumn_(usersSheet, 'Active');

  // Remove the default placeholder sheet when the real structure is ready.
  const sheets = ss.getSheets();
  if (sheets.length > 1) {
    const placeholder = ss.getSheetByName('Sheet1');
    if (placeholder) {
      ss.deleteSheet(placeholder);
    }
  }

  // Optionally create a hidden metadata sheet
  const meta = ss.getSheetByName('_meta') || ss.insertSheet('_meta');
  meta.clear();
  meta.getRange(1,1,4,2).setValues([['createdAt', new Date().toISOString()], ['createdBy', CONFIG.FROM_EMAIL], ['system','MajengoSmart'], ['version','1']]);
}

/**
 * Append a repair request into the Repairs sheet.
 */
function submitRepairRequest(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Repairs');
    const requestId = params.requestId || 'REQ-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss');
    sheet.appendRow([
      requestId,
      params.tenantName || '',
      params.tenantEmail || '',
      params.property || '',
      params.unit || '',
      params.issueDescription || '',
      params.priority || 'Medium',
      new Date(),
      'Pending',
      '',
      '',
      '',
      params.notes || ''
    ]);

    return createResponse(true, 'Repair request submitted', { requestId: requestId, sheetUrl: ss.getUrl() });
  } catch (error) {
    return createResponse(false, 'Failed to submit repair request: ' + error.toString());
  }
}

/**
 * Review a repair request and mark it approved or rejected.
 * params: { spreadsheetId?, requestId, status, actionBy, notes? }
 */
function reviewRepairRequest(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Repairs');
    const data = sheet.getDataRange().getValues();
    const requestId = params.requestId;
    const targetStatus = (params.status || '').toString();

    if (!requestId) return createResponse(false, 'requestId is required');
    if (!targetStatus) return createResponse(false, 'status is required');

    let updated = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(requestId)) {
        sheet.getRange(i + 1, 9).setValue(targetStatus);
        sheet.getRange(i + 1, 10).setValue(params.actionBy || 'Management');
        sheet.getRange(i + 1, 11).setValue(params.actionBy || 'Management');
        sheet.getRange(i + 1, 12).setValue(new Date());
        sheet.getRange(i + 1, 13).setValue(params.notes || '');
        updated = true;
        break;
      }
    }

    if (!updated) return createResponse(false, 'Repair request not found: ' + requestId);
    return createResponse(true, 'Repair request reviewed', { requestId: requestId, status: targetStatus });
  } catch (error) {
    return createResponse(false, 'Failed to review repair request: ' + error.toString());
  }
}

// ====== PUSH / PULL IMPLEMENTATIONS ======
function pushAll(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    // Push tenants
    if (params.payload && params.payload.tenants) pushArrayToSheet(ss, 'Tenants', params.payload.tenants, ['TenantID','FullName','IDNumber','Phone','Email','Property','Unit','MoveInDate','MonthlyRent','Status','EmergencyContact']);
    // Push rent payments
    if (params.payload && params.payload.rentPayments) pushArrayToSheet(ss, 'RentPayments', params.payload.rentPayments, ['PaymentID','TenantName','Unit','Amount','Method','Reference','Date','Status','RecordedBy','Notes']);
    // Push invoices
    if (params.payload && params.payload.invoices) pushArrayToSheet(ss, 'Invoices', params.payload.invoices, ['InvoiceID','TenantName','TenantEmail','Property','Unit','Amount','IssueDate','DueDate','Status','SentAt','PaymentRef','Notes']);
    // Push accounts
    if (params.payload && params.payload.accounts) pushArrayToSheet(ss, 'ChartOfAccounts', params.payload.accounts, ['Code','Name','Type','Category']);
    // Push leases
    if (params.payload && params.payload.leases) pushArrayToSheet(ss, 'Leases', params.payload.leases, ['LeaseID','Tenant','Unit','Start','End','Rent','Deposit','Status']);

    return createResponse(true, 'All data pushed');
  } catch (e) {
    return createResponse(false, 'pushAll failed: ' + e.toString());
  }
}

function pushCustomers(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    if (!params.payload || !params.payload.tenants) return createResponse(false, 'No tenants payload');
    pushArrayToSheet(ss, 'Tenants', params.payload.tenants, ['TenantID','FullName','IDNumber','Phone','Email','Property','Unit','MoveInDate','MonthlyRent','Status','EmergencyContact']);
    return createResponse(true, 'Tenants pushed');
  } catch (e) { return createResponse(false, 'pushCustomers failed: ' + e.toString()); }
}

function pushInvoices(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    if (!params.payload || !params.payload.invoices) return createResponse(false, 'No invoices payload');
    pushArrayToSheet(ss, 'Invoices', params.payload.invoices, ['InvoiceID','TenantName','TenantEmail','Property','Unit','Amount','IssueDate','DueDate','Status','SentAt','PaymentRef','Notes']);
    return createResponse(true, 'Invoices pushed');
  } catch (e) { return createResponse(false, 'pushInvoices failed: ' + e.toString()); }
}

function pushPayments(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    if (!params.payload || !params.payload.payments) return createResponse(false, 'No payments payload');
    pushArrayToSheet(ss, 'RentPayments', params.payload.payments, ['PaymentID','TenantName','Unit','Amount','Method','Reference','Date','Status','RecordedBy','Notes']);
    return createResponse(true, 'Payments pushed');
  } catch (e) { return createResponse(false, 'pushPayments failed: ' + e.toString()); }
}

function pushAccounts(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    if (!params.payload || !params.payload.accounts) return createResponse(false, 'No accounts payload');
    pushArrayToSheet(ss, 'ChartOfAccounts', params.payload.accounts, ['Code','Name','Type','Category']);
    return createResponse(true, 'Accounts pushed');
  } catch (e) { return createResponse(false, 'pushAccounts failed: ' + e.toString()); }
}

function pullAll(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const tenants = readSheetAsObjects(ss, 'Tenants');
    const rentPayments = readSheetAsObjects(ss, 'RentPayments');
    const invoices = readSheetAsObjects(ss, 'Invoices');
    const accounts = readSheetAsObjects(ss, 'ChartOfAccounts');
    const leases = readSheetAsObjects(ss, 'Leases');
    return createResponse(true, 'Pulled all data', { tenants: tenants, rentPayments: rentPayments, invoices: invoices, accounts: accounts, leases: leases });
  } catch (e) { return createResponse(false, 'pullAll failed: ' + e.toString()); }
}

function pullCustomers(params) {
  try { const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_()); createDefaultSheets(ss); const tenants = readSheetAsObjects(ss, 'Tenants'); return createResponse(true,'Pulled tenants',{tenants:tenants}); } catch(e){ return createResponse(false,'pullCustomers failed: '+e.toString()); }
}

function pullInvoices(params) {
  try { const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_()); createDefaultSheets(ss); const invoices = readSheetAsObjects(ss, 'Invoices'); return createResponse(true,'Pulled invoices',{invoices:invoices}); } catch(e){ return createResponse(false,'pullInvoices failed: '+e.toString()); }
}

function pullPayments(params) {
  try { const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_()); createDefaultSheets(ss); const payments = readSheetAsObjects(ss, 'RentPayments'); return createResponse(true,'Pulled payments',{payments:payments}); } catch(e){ return createResponse(false,'pullPayments failed: '+e.toString()); }
}

function pullAccounts(params) {
  try { const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_()); createDefaultSheets(ss); const accounts = readSheetAsObjects(ss, 'ChartOfAccounts'); return createResponse(true,'Pulled accounts',{accounts:accounts}); } catch(e){ return createResponse(false,'pullAccounts failed: '+e.toString()); }
}

function importMonthlyInvoices(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const csv = params.csv || '';
    if(!csv) return createResponse(false,'No CSV provided');
    const lines = csv.split(/\r?\n/).map(l=>l.trim()).filter(l=>l.length>0);
    const header = lines.shift().split(',').map(h=>h.trim());
    const sheet = ss.getSheetByName('Invoices');
    for(let i=0;i<lines.length;i++){
      const cols = parseCsvLine(lines[i]);
      const row = header.map((h,idx)=> cols[idx]||'');
      sheet.appendRow(row);
    }
    return createResponse(true,'Imported monthly invoices', { imported: lines.length });
  } catch(e) { return createResponse(false,'importMonthlyInvoices failed: '+e.toString()); }
}

// ====== SHEET HELPERS ======
function pushArrayToSheet(ss, sheetName, arr, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if(!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();
  if(headers && headers.length) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  if(!arr || !arr.length) return;
  const rows = arr.map(obj => headers.map(h => (obj[h] !== undefined ? obj[h] : obj[h.replace(/[^A-Za-z0-9]/g,'')] || obj[h.toLowerCase()] || '')) );
  sheet.getRange(2,1,rows.length, headers.length).setValues(rows);
}

function readSheetAsObjects(ss, sheetName){
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0].map(h=>String(h).trim());
  const out = [];
  for(let i=1;i<data.length;i++){
    const row = data[i];
    const obj = {};
    for(let c=0;c<headers.length;c++) obj[headers[c]] = row[c];
    out.push(obj);
  }
  return out;
}

function parseCsvLine(line){
  return line.split(',').map(s=>s.trim());
}

function getDefaultSpreadsheetId_() {
  return '1wuBZm0TlAfou3uVQFKG-ZK9i8hMB34EF8G3wea-Wo0k';
}

/**
 * Convenience runner for Apps Script editor.
 * If you paste the spreadsheet ID into DEFAULT_INIT_SPREADSHEET_ID, this will initialize it directly.
 */
function initiateSheet() {
  const DEFAULT_INIT_SPREADSHEET_ID = '1wuBZm0TlAfou3uVQFKG-ZK9i8hMB34EF8G3wea-Wo0k';
  const ss = SpreadsheetApp.openById(DEFAULT_INIT_SPREADSHEET_ID);
  createDefaultSheets(ss);
  return createResponse(true, 'Spreadsheet initialized', { spreadsheetId: ss.getId(), url: ss.getUrl() });
}

function doGet(e) {
  return ContentService.createTextOutput('MajengoSmart Pro Backend is running. Use POST requests.');
}

// ====== EMAIL FUNCTIONS ======

/**
 * Send invoice email to tenant
 */
function sendInvoiceEmail(params) {
  try {
    const {
      to,
      tenantName,
      propertyUnit,
      amount,
      dueDate,
      subject,
      template
    } = params;
    
    // Replace template variables
    let emailBody = template || getDefaultInvoiceTemplate();
    emailBody = emailBody.replace(/{{tenant_name}}/g, tenantName);
    emailBody = emailBody.replace(/{{property_unit}}/g, propertyUnit);
    emailBody = emailBody.replace(/{{amount}}/g, formatCurrency(amount));
    emailBody = emailBody.replace(/{{due_date}}/g, dueDate);
    emailBody = emailBody.replace(/{{company_name}}/g, CONFIG.COMPANY_NAME);
    
    // Send email
    const emailSubject = subject || `Rent Invoice - ${tenantName} - ${propertyUnit}`;
    
    MailApp.sendEmail({
      to: to,
      subject: emailSubject,
      body: emailBody,
      name: CONFIG.FROM_NAME
    });
    
    // Log the email
    logEmail({
      to: to,
      subject: emailSubject,
      status: 'sent',
      timestamp: new Date()
    });
    
    return createResponse(true, 'Email sent successfully', {
      recipient: to,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    return createResponse(false, 'Failed to send email: ' + error.toString());
  }
}

/**
 * Send generic email
 */
function sendEmail(params) {
  try {
    const { to, subject, body, from, fromName, fromEmail } = params;

    // Try to get company name/email from spreadsheet _meta or fallback to CONFIG
    let companyName = fromName || CONFIG.FROM_NAME;
    let companyEmail = fromEmail || from || CONFIG.FROM_EMAIL;

    // Attempt to read company name from _meta sheet
    try {
      const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
      const metaSheet = ss.getSheetByName('_meta');
      if (metaSheet) {
        const metaData = metaSheet.getDataRange().getValues();
        metaData.forEach(function(row) {
          if (String(row[0]).trim() === 'companyName' && row[1]) companyName = String(row[1]).trim();
          if (String(row[0]).trim() === 'companyEmail' && row[1]) companyEmail = String(row[1]).trim();
        });
      }
    } catch(metaErr) { /* use defaults */ }

    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      name: companyName,
      replyTo: companyEmail
    });

    return createResponse(true, 'Email sent successfully');

  } catch (error) {
    return createResponse(false, 'Failed to send email: ' + error.toString());
  }
}

/**
 * Get default invoice email template
 */
function getDefaultInvoiceTemplate() {
  return `Dear {{tenant_name}},

Your rent invoice for {{property_unit}} is now due.

Amount Due: KES {{amount}}
Due Date: {{due_date}}

Please make payment via:
- M-Pesa Paybill: 174379
- Equity Bank A/C: 0123456789
- KCB Bank A/C: 9876543210

Thank you for your prompt payment.

Best regards,
{{company_name}}
+254 712 345 678
info@majengosmart.co.ke`;
}

// ====== GOOGLE DRIVE STORAGE FUNCTIONS ======

/**
 * Create storage folder structure in Google Drive
 */
function createStorageFolders(params) {
  try {
    // Get the configured root folder (TENANT MANAGEMENT SYSTEM)
    let mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
    
    // Create subfolders inside the root folder
    const subfolders = [
      'Active Tenants',
      'Archived Tenants',
      'Lease Agreements',
      'Financial Records',
      'Property Documents',
      'Maintenance Records',
      'Backups',
      'Email Logs'
    ];
    
    const created = [];
    subfolders.forEach(name => {
      const folder = getOrCreateFolder(name, mainFolder);
      created.push({
        name: name,
        id: folder.getId(),
        url: folder.getUrl()
      });
    });
    
    return createResponse(true, 'Folders created successfully in TENANT MANAGEMENT SYSTEM', {
      mainFolderId: mainFolder.getId(),
      mainFolderUrl: mainFolder.getUrl(),
      subfolders: created
    });
    
  } catch (error) {
    return createResponse(false, 'Failed to create folders: ' + error.toString());
  }
}

/**
 * Archive tenant data to Google Drive
 */
function archiveTenant(params) {
  try {
    const {
      tenantId,
      tenantName,
      tenantData,
      leaseData,
      paymentHistory,
      maintenanceHistory
    } = params;
    
    // Get archive folder from the configured root folder
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
    const archiveFolder = getOrCreateFolder('Archived Tenants', mainFolder);
    
    // Create tenant-specific folder
    const tenantFolderName = `${tenantName} - ${tenantId}`;
    const tenantFolder = getOrCreateFolder(tenantFolderName, archiveFolder);
    
    // Create JSON files for each data type
    const files = [];
    
    if (tenantData) {
      const file = tenantFolder.createFile('tenant_profile.json', JSON.stringify(tenantData, null, 2));
      files.push(file.getName());
    }
    
    if (leaseData) {
      const file = tenantFolder.createFile('lease_agreements.json', JSON.stringify(leaseData, null, 2));
      files.push(file.getName());
    }
    
    if (paymentHistory) {
      const file = tenantFolder.createFile('payment_history.json', JSON.stringify(paymentHistory, null, 2));
      files.push(file.getName());
    }
    
    if (maintenanceHistory) {
      const file = tenantFolder.createFile('maintenance_history.json', JSON.stringify(maintenanceHistory, null, 2));
      files.push(file.getName());
    }
    
    // Create archive summary
    const summary = {
      tenantId: tenantId,
      tenantName: tenantName,
      archivedDate: new Date().toISOString(),
      filesArchived: files
    };
    
    tenantFolder.createFile('archive_summary.json', JSON.stringify(summary, null, 2));
    
    return createResponse(true, 'Tenant archived successfully', {
      folderId: tenantFolder.getId(),
      folderUrl: tenantFolder.getUrl(),
      filesCreated: files.length + 1
    });
    
  } catch (error) {
    return createResponse(false, 'Failed to archive tenant: ' + error.toString());
  }
}

/**
 * Backup all system data
 */
function backupData(params) {
  try {
    const { data, backupName } = params;
    
    // Get backup folder from the configured root folder
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
    const backupFolder = getOrCreateFolder('Backups', mainFolder);
    
    // Create backup file
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm-ss');
    const fileName = backupName || `backup_${timestamp}.json`;
    
    const file = backupFolder.createFile(fileName, JSON.stringify(data, null, 2));
    
    return createResponse(true, 'Backup created successfully', {
      fileId: file.getId(),
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      timestamp: timestamp
    });
    
  } catch (error) {
    return createResponse(false, 'Failed to create backup: ' + error.toString());
  }
}

/**
 * Restore data from backup
 */
function restoreData(params) {
  try {
    const { fileId } = params;
    
    // Get the backup file
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    const data = JSON.parse(content);
    
    return createResponse(true, 'Backup restored successfully', {
      data: data,
      fileName: file.getName()
    });
    
  } catch (error) {
    return createResponse(false, 'Failed to restore backup: ' + error.toString());
  }
}

// ====== HELPER FUNCTIONS ======

/**
 * Get or create a folder
 */
function getOrCreateFolder(folderName, parentFolder) {
  const parent = parentFolder || DriveApp.getRootFolder();
  const folders = parent.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parent.createFolder(folderName);
  }
}

/**
 * Log email to spreadsheet
 */
function logEmail(emailData) {
  try {
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
    const logFolder = getOrCreateFolder('Email Logs', mainFolder);
    
    // Get or create log spreadsheet
    const files = logFolder.getFilesByName('email_log.csv');
    let content = '';
    
    if (files.hasNext()) {
      const file = files.next();
      content = file.getBlob().getDataAsString();
    } else {
      content = 'Timestamp,To,Subject,Status\n';
    }
    
    // Append new log entry
    const timestamp = Utilities.formatDate(emailData.timestamp, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const row = `${timestamp},"${emailData.to}","${emailData.subject}","${emailData.status}"\n`;
    content += row;
    
    // Save log
    const file = logFolder.createFile('email_log.csv', content, MimeType.CSV);
    
  } catch (error) {
    Logger.log('Failed to log email: ' + error.toString());
  }
}

/**
 * Format currency
 */
function formatCurrency(amount) {
  return 'KES ' + Number(amount).toLocaleString('en-KE', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Create standardized response
 */
function createResponse(success, message, data) {
  const response = {
    success: success,
    message: message,
    data: data || {},
    timestamp: new Date().toISOString()
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== TEST FUNCTIONS ======

/**
 * Test email sending
 */
function testEmail() {
  const params = {
    action: 'send_invoice',
    to: 'test@example.com',
    tenantName: 'John Kamau',
    propertyUnit: 'A1',
    amount: 20000,
    dueDate: '2026-06-01',
    subject: 'Test Invoice',
    template: getDefaultInvoiceTemplate()
  };
  
  const result = sendInvoiceEmail(params);
  Logger.log(result.getContent());
}

/**
 * Test folder creation
 */
function testFolderCreation() {
  const result = createStorageFolders({});
  Logger.log(result.getContent());
}

/**
 * Test tenant archiving
 */
function testTenantArchive() {
  const params = {
    tenantId: 'T001',
    tenantName: 'John Kamau',
    tenantData: {
      name: 'John Kamau',
      email: 'john@email.com',
      phone: '+254 712 345 678',
      unit: 'A1'
    },
    leaseData: {
      start: '2025-01-15',
      end: '2026-01-14',
      rent: 20000
    },
    paymentHistory: [
      {date: '2026-01-01', amount: 20000, method: 'M-Pesa'},
      {date: '2026-02-01', amount: 20000, method: 'M-Pesa'}
    ]
  };
  
  const result = archiveTenant(params);
  Logger.log(result.getContent());
}

// ====== TEAMS & USERS MANAGEMENT ======

function createTeam(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Teams');
    ensureStatusColumn_(sheet, 'Active');
    const teamId = 'TEAM-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss');
    const permissions = Array.isArray(params.permissions) ? params.permissions.join(',') : (params.permissions || '');
    sheet.appendRow([teamId, params.teamName || ('Team ' + teamId), permissions, new Date(), params.notes || '', 'Active']);
    return createResponse(true, 'Team created', { teamId: teamId });
  } catch (e) {
    return createResponse(false, 'createTeam failed: ' + e.toString());
  }
}

function listTeams(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const teams = readSheetAsObjects(ss, 'Teams');
    return createResponse(true, 'Teams listed', { teams: teams });
  } catch (e) {
    return createResponse(false, 'listTeams failed: ' + e.toString());
  }
}

function createUserAccount(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Users');
    ensureStatusColumn_(sheet, 'Active');
    const userId = 'U-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss');
    const email = params.email || '';
    const fullName = params.fullName || '';
    const team = params.team || '';
    const role = params.role || 'Member';
    const password = params.password || generateTempPassword();
    sheet.appendRow([userId, email, fullName, team, role, password, new Date(), 'Active']);

    // ── Read company name & email from _meta sheet (set by Company Details panel) ──
    let companyName = params.companyName || CONFIG.COMPANY_NAME;
    let companyEmail = params.companyEmail || CONFIG.FROM_EMAIL;
    try {
      const metaSheet = ss.getSheetByName('_meta');
      if (metaSheet) {
        const metaData = metaSheet.getDataRange().getValues();
        metaData.forEach(function(row) {
          if (String(row[0]).trim() === 'companyName' && row[1]) companyName = String(row[1]).trim();
          if (String(row[0]).trim() === 'companyEmail' && row[1]) companyEmail = String(row[1]).trim();
        });
      }
    } catch(metaErr) { /* use defaults */ }

    // ── Send welcome email immediately ──
    if (email) {
      const subject = `Your ${companyName} Account Has Been Created`;
      const body =
        `Dear ${fullName || 'User'},\n\n` +
        `Welcome to ${companyName}!\n\n` +
        `An account has been created for you. Below are your login credentials:\n\n` +
        `  Username / Email  : ${email}\n` +
        `  Temporary Password: ${password}\n` +
        `  Role              : ${role}\n\n` +
        `Please log in and change your password immediately after your first login.\n\n` +
        `Regards,\n${companyName} Administration\n${companyEmail}`;
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body,
          name: companyName,
          replyTo: companyEmail
        });
        Logger.log('Welcome email sent to ' + email + ' from ' + companyEmail);
      } catch (ee) {
        Logger.log('Failed to send invite email: ' + ee.toString());
      }
    }

    return createResponse(true, 'User account created', { userId: userId, password: password });
  } catch (e) {
    return createResponse(false, 'createUserAccount failed: ' + e.toString());
  }
}

function listUsers(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const users = readSheetAsObjects(ss, 'Users');
    return createResponse(true, 'Users listed', { users: users });
  } catch (e) {
    return createResponse(false, 'listUsers failed: ' + e.toString());
  }
}

function setTeamPermissions(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Teams');
    const data = sheet.getDataRange().getValues();
    const teamId = params.teamId;
    if (!teamId) return createResponse(false, 'teamId is required');
    let updated = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(teamId)) {
        sheet.getRange(i + 1, 3).setValue(Array.isArray(params.permissions) ? params.permissions.join(',') : (params.permissions || ''));
        updated = true;
        break;
      }
    }
    if (!updated) return createResponse(false, 'Team not found: ' + teamId);
    return createResponse(true, 'Permissions updated');
  } catch (e) { return createResponse(false, 'setTeamPermissions failed: ' + e.toString()); }
}

function setUserStatus(params, status) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Users');
    const userId = params.userId || params.userID || params.id;
    if (!userId) return createResponse(false, 'userId is required');
    ensureStatusColumn_(sheet, 'Active');
    return updateRowStatus_(sheet, userId, ['UserID', 'UserId', 'ID'], 'Status', status, 'user', 'userId');
  } catch (e) {
    return createResponse(false, 'setUserStatus failed: ' + e.toString());
  }
}

function setTeamStatus(params, status) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Teams');
    const teamId = params.teamId || params.teamID || params.id;
    if (!teamId) return createResponse(false, 'teamId is required');
    ensureStatusColumn_(sheet, 'Active');
    return updateRowStatus_(sheet, teamId, ['TeamID', 'TeamId', 'ID'], 'Status', status, 'team', 'teamId');
  } catch (e) {
    return createResponse(false, 'setTeamStatus failed: ' + e.toString());
  }
}

function generateTempPassword() {
  // Simple temporary password generator
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let p = '';
  for (let i = 0; i < 10; i++) p += chars.charAt(Math.floor(Math.random() * chars.length));
  return p + '1a!';
}

function ensureStatusColumn_(sheet, defaultStatus) {
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return null;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header).trim();
  });
  const statusIndex = headers.findIndex(function(header) {
    return header.toLowerCase() === 'status';
  });

  if (statusIndex !== -1) {
    return statusIndex + 1;
  }

  const newColumn = lastColumn + 1;
  sheet.getRange(1, newColumn).setValue('Status');
  if (lastRow > 1) {
    const values = [];
    for (let i = 2; i <= lastRow; i++) {
      values.push([defaultStatus || 'Active']);
    }
    sheet.getRange(2, newColumn, values.length, 1).setValues(values);
  }
  return newColumn;
}

function updateRowStatus_(sheet, recordId, idHeaders, statusHeader, statusValue, entityLabel, entityIdLabel) {
  const data = sheet.getDataRange().getValues();
  if (!data.length) return createResponse(false, entityLabel + ' sheet is empty');

  const headers = data[0].map(function(header) {
    return String(header).trim();
  });
  const idIndex = findHeaderIndex_(headers, idHeaders);
  if (idIndex === -1) return createResponse(false, entityLabel + ' ID column not found');

  let statusIndex = findHeaderIndex_(headers, [statusHeader]);
  if (statusIndex === -1) {
    statusIndex = ensureStatusColumn_(sheet, 'Active') - 1;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]).trim() === String(recordId).trim()) {
      sheet.getRange(i + 1, statusIndex + 1).setValue(statusValue);
      return createResponse(true, entityLabel + ' status updated', {
        [entityIdLabel]: recordId,
        status: statusValue
      });
    }
  }

  return createResponse(false, entityLabel + ' not found: ' + recordId);
}

function findHeaderIndex_(headers, candidates) {
  const normalizedCandidates = candidates.map(function(candidate) {
    return String(candidate).trim().toLowerCase();
  });
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).trim().toLowerCase();
    if (normalizedCandidates.indexOf(header) !== -1) {
      return i;
    }
  }
  return -1;
}

/**
 * Test backup creation
 */
function testBackup() {
  const params = {
    data: {
      tenants: [{name: 'John Kamau', unit: 'A1'}],
      properties: [{name: 'Majengo Plaza', units: 15}]
    },
    backupName: 'test_backup.json'
  };
  
  const result = backupData(params);
  Logger.log(result.getContent());
}

/**
 * Schedule auto-send settings (saved to a Scheduler sheet).
 * params: { enabled: true/false, dayOfMonth: number, spreadsheetId?: string }
 */
function scheduleAutoSend(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    let sheet = ss.getSheetByName('Scheduler');
    if(!sheet) sheet = ss.insertSheet('Scheduler');
    // Ensure header
    sheet.getRange(1,1,1,4).setValues([['enabled','dayOfMonth','lastRun','updatedAt']]);
    // For simplicity only one scheduler row is used (row 2)
    sheet.getRange(2,1,1,4).setValues([[params.enabled ? 'true' : 'false', params.dayOfMonth || '15', params.lastRun || '', new Date().toISOString()]]);
    return createResponse(true, 'Auto-send scheduler updated', { enabled: params.enabled, dayOfMonth: params.dayOfMonth });
  } catch (error) {
    return createResponse(false, 'Failed to update scheduler: ' + error.toString());
  }
}

/**
 * Generate monthly billing invoices from tenant records.
 * params: { spreadsheetId?: string, month?: number, year?: number, property?: string }
 */
function generateBilling(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);

    const tenants = readSheetAsObjects(ss, 'Tenants');
    const invoicesSheet = ss.getSheetByName('Invoices');
    if(!invoicesSheet) return createResponse(false, 'Invoices sheet not found');

    const month = Number(params.month) || (new Date()).getMonth() + 1;
    const year = Number(params.year) || (new Date()).getFullYear();
    const propertyFilter = String(params.property || 'all');
    const issueDate = new Date(year, month - 1, 1);
    const dayOfMonth = Math.min(Math.max(Number(params.dayOfMonth) || 15, 1), 28);
    const dueDate = new Date(year, month - 1, dayOfMonth);
    const issueDateText = Utilities.formatDate(issueDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const dueDateText = Utilities.formatDate(dueDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');

    const existing = readSheetAsObjects(ss, 'Invoices');
    const existingKeys = {};
    existing.forEach(inv => {
      const existingIssue = new Date(inv.IssueDate || inv.issueDate || inv.Date || inv.date || '');
      const key = [String(inv.TenantName || inv.tenant || '').trim().toLowerCase(), String(inv.Property || inv.property || '').trim().toLowerCase(), existingIssue.getMonth() + 1, existingIssue.getFullYear()].join('|');
      existingKeys[key] = true;
    });

    const rows = [];
    let generated = 0;
    tenants.forEach(tenant => {
      const tenantName = String(tenant.FullName || tenant.TenantName || tenant.name || '').trim();
      const tenantProperty = String(tenant.Property || tenant.property || '').trim();
      const tenantStatus = String(tenant.Status || tenant.status || 'active').toLowerCase();
      if(!tenantName) return;
      if(propertyFilter !== 'all' && tenantProperty !== propertyFilter) return;
      if(tenantStatus === 'archived') return;

      const key = [tenantName.toLowerCase(), tenantProperty.toLowerCase(), month, year].join('|');
      if(existingKeys[key]) return;

      const invoiceId = 'INV-' + year + String(month).padStart(2, '0') + '-' + String(tenant.Unit || tenant.unit || tenantName).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
      rows.push([
        invoiceId,
        tenantName,
        tenant.Email || tenant.email || '',
        tenantProperty,
        tenant.Unit || tenant.unit || '',
        Number(tenant.MonthlyRent || tenant.rent || 0),
        issueDateText,
        dueDateText,
        'Draft',
        '',
        'BILL-' + year + String(month).padStart(2, '0'),
        'Generated by Run Billing'
      ]);
      existingKeys[key] = true;
      generated++;
    });

    if(!rows.length) {
      return createResponse(true, 'No billing records generated', { generated: 0, month: month, year: year, property: propertyFilter });
    }

    const startRow = invoicesSheet.getLastRow() + 1;
    invoicesSheet.getRange(startRow, 1, rows.length, 12).setValues(rows);

    return createResponse(true, 'Billing generated', { generated: generated, month: month, year: year, property: propertyFilter, issueDate: issueDateText, dueDate: dueDateText });
  } catch (error) {
    return createResponse(false, 'generateBilling failed: ' + error.toString());
  }
}

/**
 * Time-driven runner: reads Scheduler and sends invoices for the current month when due.
 * Intended to be attached as an Apps Script time-driven trigger (daily).
 */
function runScheduledSends(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    const sheet = ss.getSheetByName('Scheduler');
    if(!sheet) return createResponse(false, 'No scheduler configured');
    const row = sheet.getRange(2,1,1,4).getValues()[0];
    const enabled = String(row[0]) === 'true';
    const day = Number(row[1]) || 15;
    const lastRun = row[2] || '';
    const today = new Date();
    const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd');

    if(!enabled) return createResponse(true, 'Scheduler disabled');
    if(today.getDate() !== day) return createResponse(true, 'Not scheduled for today');
    if(lastRun === todayStr) return createResponse(true, 'Already ran today');

    // Send invoices with Status != 'Sent' for the current month
    const invoices = readSheetAsObjects(ss, 'Invoices');
    const sent = [];
    invoices.forEach(inv => {
      const issue = new Date(inv.IssueDate || inv.IssueDate || inv.issueDate || inv.Issue_Date || '');
      const due = inv.DueDate || inv.DueDate || inv.dueDate || '';
      // Only send invoices for the current month/year
      if(isNaN(issue)) return;
      if(issue.getMonth() === today.getMonth() && issue.getFullYear() === today.getFullYear()) {
        if(String(inv.Status || inv.status || '').toLowerCase() !== 'sent') {
          const template = getDefaultInvoiceTemplate();
          const p = { action: 'send_invoice', to: inv.TenantEmail || inv.tenantEmail || '', tenantName: inv.TenantName || inv.tenant || '', propertyUnit: inv.Unit || inv.unit || '', amount: inv.Amount || inv.amount || 0, dueDate: due, subject: `Rent Invoice - ${inv.TenantName || ''}`, template: template };
          try {
            sendInvoiceEmail(p);
            sent.push(inv.InvoiceID || inv.InvoiceId || inv.id || '');
          } catch(e) {
            Logger.log('Failed to send invoice ' + (inv.InvoiceID||'') + ': ' + e.toString());
          }
        }
      }
    });

    // Update lastRun
    sheet.getRange(2,3).setValue(todayStr);

    return createResponse(true, 'Scheduled send completed', { sentCount: sent.length, sent: sent });
  } catch (error) {
    return createResponse(false, 'runScheduledSends failed: ' + error.toString());
  }
}

/**
 * Simple Drive connectivity test: verifies configured DRIVE_ROOT_FOLDER_ID exists and returns info
 */
function testDriveConnection(params) {
  try {
    const id = params.folderId || CONFIG.DRIVE_ROOT_FOLDER_ID;
    const folder = DriveApp.getFolderById(id);
    return createResponse(true, 'Drive connection OK', { id: folder.getId(), url: folder.getUrl(), name: folder.getName() });
  } catch (error) {
    return createResponse(false, 'Drive connection failed: ' + error.toString());
  }
}

/**
 * Manual billing trigger: callable from both webapp and backend.
 * This allows billing to be run on-demand from the backend without waiting for webapp interaction.
 * params: { spreadsheetId?, month?, year?, property?, dayOfMonth? }
 */
function runBillingManual(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);

    const month = Number(params.month) || (new Date()).getMonth() + 1;
    const year = Number(params.year) || (new Date()).getFullYear();
    const property = String(params.property || 'all');
    const dayOfMonth = Math.min(Math.max(Number(params.dayOfMonth) || 15, 1), 28);

    // Call generateBilling with the provided parameters
    const billingParams = {
      spreadsheetId: params.spreadsheetId || getDefaultSpreadsheetId_(),
      month: month,
      year: year,
      property: property,
      dayOfMonth: dayOfMonth
    };

    const result = generateBilling(billingParams);
    
    // Parse response to extract data
    const response = JSON.parse(result.getContent());
    
    return createResponse(true, 'Manual billing completed', response.data || {});
  } catch (error) {
    return createResponse(false, 'runBillingManual failed: ' + error.toString());
  }
}

/**
 * Convenience function: run billing for current month on demand from backend.
 * Useful for time-driven triggers or standalone execution.
 * Can be called as: runBillingCurrentMonth() from Apps Script editor.
 */
function runBillingCurrentMonth() {
  const now = new Date();
  const params = {
    spreadsheetId: getDefaultSpreadsheetId_(),
    month: now.getMonth() + 1,
    year: now.getFullYear(),
    property: 'all'
  };
  const response = runBillingManual(params);
  Logger.log(response.getContent());
  return response;
}

/**
 * Schedule a billing run for a specific date/time.
 * This can be triggered via time-driven trigger: e.g., every 1st of the month at 9 AM
 * To use: Set up a time-driven trigger in Apps Script editor → Triggers → "runBillingScheduled"
 */
function runBillingScheduled() {
  try {
    const now = new Date();
    const ss = SpreadsheetApp.openById(getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    // Check if scheduler is enabled and it's the right day
    const schedulerSheet = ss.getSheetByName('Scheduler');
    if(!schedulerSheet) {
      Logger.log('Scheduler sheet not found');
      return;
    }
    
    const schedulerRow = schedulerSheet.getRange(2,1,1,4).getValues()[0];
    const enabled = String(schedulerRow[0]) === 'true';
    const billingDay = Number(schedulerRow[1]) || 1;
    
    if(!enabled) {
      Logger.log('Billing scheduler is disabled');
      return;
    }
    
    if(now.getDate() !== billingDay) {
      Logger.log(`Not billing day today (expected ${billingDay}, got ${now.getDate()})`);
      return;
    }
    
    // Run billing for current month
    const result = runBillingCurrentMonth();
    Logger.log('Scheduled billing completed: ' + result.getContent());
  } catch (error) {
    Logger.log('runBillingScheduled failed: ' + error.toString());
  }
}

// ====== CHART OF ACCOUNTS FUNCTIONS ======

/**
 * Initialize default Chart of Accounts with standard accounting structure
 */
function initializeChartOfAccounts(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    const accounts = [
      // ASSETS (1000-1999)
      {code: '1000', name: 'Cash in Bank', category: 'Assets', type: 'Current Asset', description: 'Bank accounts'},
      {code: '1010', name: 'Petty Cash', category: 'Assets', type: 'Current Asset', description: 'Cash on hand'},
      {code: '1100', name: 'Tenant Deposits', category: 'Assets', type: 'Current Asset', description: 'Deposit accounts'},
      {code: '1200', name: 'Accounts Receivable', category: 'Assets', type: 'Current Asset', description: 'Outstanding rent'},
      {code: '1500', name: 'Property Equipment', category: 'Assets', type: 'Fixed Asset', description: 'Furniture and equipment'},
      {code: '1600', name: 'Accumulated Depreciation', category: 'Assets', type: 'Fixed Asset', description: 'Equipment depreciation'},
      
      // LIABILITIES (2000-2999)
      {code: '2000', name: 'Accounts Payable', category: 'Liabilities', type: 'Current Liability', description: 'Vendor payments due'},
      {code: '2100', name: 'Rent Payable', category: 'Liabilities', type: 'Current Liability', description: 'Property rent owed'},
      {code: '2200', name: 'Salaries Payable', category: 'Liabilities', type: 'Current Liability', description: 'Staff wages due'},
      {code: '2300', name: 'Maintenance Payable', category: 'Liabilities', type: 'Current Liability', description: 'Maintenance work due'},
      
      // EQUITY (3000-3999)
      {code: '3000', name: 'Owner Capital', category: 'Equity', type: 'Equity', description: 'Owner investment'},
      {code: '3100', name: 'Retained Earnings', category: 'Equity', type: 'Equity', description: 'Prior year earnings'},
      
      // INCOME (4000-4999)
      {code: '4000', name: 'Rental Income', category: 'Income', type: 'Revenue', description: 'Rent from tenants'},
      {code: '4100', name: 'Service Charges', category: 'Income', type: 'Revenue', description: 'Water, electricity charges'},
      {code: '4200', name: 'Parking Income', category: 'Income', type: 'Revenue', description: 'Parking fees'},
      {code: '4300', name: 'Late Payment Fees', category: 'Income', type: 'Revenue', description: 'Penalties'},
      {code: '4400', name: 'Other Income', category: 'Income', type: 'Revenue', description: 'Miscellaneous income'},
      
      // EXPENSES (5000-5999)
      {code: '5000', name: 'Maintenance & Repairs', category: 'Expenses', type: 'Operating Expense', description: 'Building maintenance'},
      {code: '5100', name: 'Utilities', category: 'Expenses', type: 'Operating Expense', description: 'Water and electricity'},
      {code: '5200', name: 'Cleaning Services', category: 'Expenses', type: 'Operating Expense', description: 'Janitorial services'},
      {code: '5300', name: 'Security', category: 'Expenses', type: 'Operating Expense', description: 'Security services'},
      {code: '5400', name: 'Staff Salaries', category: 'Expenses', type: 'Operating Expense', description: 'Employee salaries'},
      {code: '5500', name: 'Insurance', category: 'Expenses', type: 'Operating Expense', description: 'Property insurance'},
      {code: '5600', name: 'Taxes - KRA MRI', category: 'Expenses', type: 'Operating Expense', description: 'Property tax'},
      {code: '5700', name: 'Advertising', category: 'Expenses', type: 'Operating Expense', description: 'Marketing and ads'},
      {code: '5800', name: 'Professional Fees', category: 'Expenses', type: 'Operating Expense', description: 'Legal and accounting'},
      {code: '5900', name: 'Miscellaneous Expenses', category: 'Expenses', type: 'Operating Expense', description: 'Other expenses'}
    ];
    
    // Write to ChartOfAccounts sheet
    let sheet = ss.getSheetByName('ChartOfAccounts');
    if(!sheet) sheet = ss.insertSheet('ChartOfAccounts');
    sheet.clear();
    
    const headers = ['Code', 'Name', 'Category', 'Type', 'Description', 'Balance', 'LastUpdated'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    
    const rows = accounts.map(acc => [acc.code, acc.name, acc.category, acc.type, acc.description, 0, new Date().toISOString()]);
    sheet.getRange(2,1,rows.length, headers.length).setValues(rows);
    
    return createResponse(true, 'Chart of Accounts initialized', {count: accounts.length, url: ss.getUrl()});
  } catch (error) {
    return createResponse(false, 'Failed to initialize COA: ' + error.toString());
  }
}

/**
 * Get Chart of Accounts with filter by category
 */
function getChartOfAccounts(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    let accounts = readSheetAsObjects(ss, 'ChartOfAccounts');
    if(!accounts || !accounts.length) {
      // Initialize if empty
      initializeChartOfAccounts(params);
      accounts = readSheetAsObjects(ss, 'ChartOfAccounts');
    }
    
    const category = params.category || 'all';
    if(category !== 'all') {
      accounts = accounts.filter(a => String(a.Category || a.category).toLowerCase() === category.toLowerCase());
    }
    
    return createResponse(true, 'Chart of Accounts retrieved', {accounts: accounts});
  } catch (error) {
    return createResponse(false, 'Failed to retrieve COA: ' + error.toString());
  }
}

// ====== FINANCIAL REPORTS FUNCTIONS ======

/**
 * Generate P&L Statement with optional comparison
 * params: {spreadsheetId?, month, year, compareMonth?, compareYear?, property?}
 */
function generateProfitLoss(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    const month = Number(params.month) || (new Date()).getMonth() + 1;
    const year = Number(params.year) || (new Date()).getFullYear();
    const compareMonth = params.compareMonth ? Number(params.compareMonth) : null;
    const compareYear = params.compareYear ? Number(params.compareYear) : null;
    
    // Get invoices for period
    const invoices = readSheetAsObjects(ss, 'Invoices');
    const expenses = readSheetAsObjects(ss, 'Invoices'); // Use same for demo, ideally from Expenses sheet
    
    const filterByPeriod = (records, m, y) => {
      return records.filter(r => {
        const date = new Date(r.IssueDate || r.issueDate || '');
        return date.getMonth() + 1 === m && date.getFullYear() === y;
      });
    };
    
    const periodA = filterByPeriod(invoices, month, year);
    const periodB = compareMonth && compareYear ? filterByPeriod(invoices, compareMonth, compareYear) : null;
    
    const totalRevenueA = periodA.reduce((s, i) => s + (Number(i.Amount || i.amount || 0)), 0);
    const totalRevenueB = periodB ? periodB.reduce((s, i) => s + (Number(i.Amount || i.amount || 0)), 0) : 0;
    
    return createResponse(true, 'P&L generated', {
      period: {month, year},
      comparison: {month: compareMonth, year: compareYear},
      revenue: {current: totalRevenueA, comparison: totalRevenueB},
      expenses: {current: 0, comparison: 0},
      netProfit: {current: totalRevenueA, comparison: totalRevenueB}
    });
  } catch (error) {
    return createResponse(false, 'generateProfitLoss failed: ' + error.toString());
  }
}

/**
 * Generate Balance Sheet with optional comparison
 */
function generateBalanceSheet(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    const month = Number(params.month) || (new Date()).getMonth() + 1;
    const year = Number(params.year) || (new Date()).getFullYear();
    
    const coa = readSheetAsObjects(ss, 'ChartOfAccounts');
    const assets = coa.filter(a => String(a.Category || a.category) === 'Assets');
    const liabilities = coa.filter(a => String(a.Category || a.category) === 'Liabilities');
    const equity = coa.filter(a => String(a.Category || a.category) === 'Equity');
    
    return createResponse(true, 'Balance Sheet generated', {
      period: {month, year},
      assets: assets,
      liabilities: liabilities,
      equity: equity,
      totalAssets: assets.reduce((s, a) => s + (Number(a.Balance || 0)), 0),
      totalLiabilities: liabilities.reduce((s, l) => s + (Number(l.Balance || 0)), 0),
      totalEquity: equity.reduce((s, e) => s + (Number(e.Balance || 0)), 0)
    });
  } catch (error) {
    return createResponse(false, 'generateBalanceSheet failed: ' + error.toString());
  }
}

/**
 * Generate Trial Balance with optional comparison
 */
function generateTrialBalance(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    const month = Number(params.month) || (new Date()).getMonth() + 1;
    const year = Number(params.year) || (new Date()).getFullYear();
    const compareMonth = params.compareMonth ? Number(params.compareMonth) : null;
    const compareYear = params.compareYear ? Number(params.compareYear) : null;
    
    const coa = readSheetAsObjects(ss, 'ChartOfAccounts');
    
    return createResponse(true, 'Trial Balance generated', {
      period: {month, year},
      comparison: {month: compareMonth, year: compareYear},
      accounts: coa.map(a => ({
        code: a.Code || a.code,
        name: a.Name || a.name,
        debit: Number(a.Balance || 0),
        credit: 0
      })),
      totalDebit: coa.reduce((s, a) => s + (Number(a.Balance || 0)), 0),
      totalCredit: 0
    });
  } catch (error) {
    return createResponse(false, 'generateTrialBalance failed: ' + error.toString());
  }
}

/**
 * Get monthly chart data for all months
 */
function getMonthlyCharts(params) {
  try {
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);
    
    const year = Number(params.year) || (new Date()).getFullYear();
    const invoices = readSheetAsObjects(ss, 'Invoices');
    
    const monthlyData = {};
    for(let m = 1; m <= 12; m++) {
      const monthlyInvoices = invoices.filter(i => {
        const date = new Date(i.IssueDate || i.issueDate || '');
        return date.getMonth() + 1 === m && date.getFullYear() === year;
      });
      monthlyData[getMonthName(m)] = monthlyInvoices.reduce((s, i) => s + (Number(i.Amount || i.amount || 0)), 0);
    }
    
    return createResponse(true, 'Monthly charts retrieved', {year, data: monthlyData});
  } catch (error) {
    return createResponse(false, 'getMonthlyCharts failed: ' + error.toString());
  }
}

function getMonthName(m) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  return months[m - 1] || 'Unknown';
}

// ======================================================================
// ====== SHEET BEAUTIFIER — formatAllSheets() ==========================
// ======================================================================
/**
 * Applies professional table formatting to every data sheet in the spreadsheet.
 * Callable via Web App: { action: 'format_all_sheets', spreadsheetId: '...' }
 * Also callable directly from the Script Editor: just run formatAllSheets({})
 */
function formatAllSheets(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(params.spreadsheetId || getDefaultSpreadsheetId_());
    createDefaultSheets(ss);

    // ── Color palette (navy / green brand) ──────────────────────────────
    const CLR = {
      headerBg:    '#14213D',   // Navy – header background
      headerFg:    '#FFFFFF',   // White – header text
      headerBorder:'#1A6B3A',   // Green – outer border accent
      bandLight:   '#FFFFFF',   // Row 1 band
      bandDark:    '#EBF4EF',   // Row 2 band (pale green)
      totalRowBg:  '#1A6B3A',   // Totals / summary row
      totalRowFg:  '#FFFFFF',
      statusActive:'#D4EDDA',   // Soft green
      statusInact: '#F8D7DA',   // Soft red
      statusPend:  '#FFF3CD',   // Soft amber
      statusSent:  '#D1ECF1',   // Soft blue
      tabActive:   '#1A6B3A',   // Sheet tab – active sheets
      tabArchive:  '#6C757D',   // Sheet tab – _meta / archive
    };

    // Sheets to style with their accent tab colours (all data sheets listed here)
    const SHEET_META = {
      'Tenants':              { tab: CLR.tabActive,  emoji: '👥', statusCol: 'Status' },
      'RentPayments':         { tab: '#C8922A',       emoji: '💰', statusCol: 'Status' },
      'Invoices':             { tab: '#2D9E57',       emoji: '🧾', statusCol: 'Status' },
      'Repairs':              { tab: '#E65C00',       emoji: '🔧', statusCol: 'Status' },
      'MaintenanceApprovals': { tab: '#8B5CF6',       emoji: '✅', statusCol: 'Status' },
      'TenantStatements':     { tab: '#0EA5E9',       emoji: '📄', statusCol: null },
      'Aging_Tenants':        { tab: '#DC2626',       emoji: '⚠️', statusCol: null },
      'Aging_Suppliers':      { tab: '#DC2626',       emoji: '⚠️', statusCol: null },
      'Teams':                { tab: '#7C3AED',       emoji: '🏢', statusCol: 'Status' },
      'Users':                { tab: '#1D4ED8',       emoji: '👤', statusCol: 'Status' },
      'ChartOfAccounts':      { tab: '#065F46',       emoji: '📊', statusCol: null },
      'Leases':               { tab: '#92400E',       emoji: '📋', statusCol: 'Status' },
    };

    const results = [];

    ss.getSheets().forEach(function(sheet) {
      const name = sheet.getName();

      // Skip hidden _meta sheet (just recolor its tab)
      if (name === '_meta') {
        sheet.setTabColor(CLR.tabArchive);
        results.push(name + ': tab recolored');
        return;
      }

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      // Skip truly empty sheets (no columns at all)
      if (lastCol < 1) {
        results.push(name + ': no columns, skipped');
        return;
      }
      // If only headers exist (lastRow=1) still format — just skip data-only steps

      // ── 1. Freeze header row ──────────────────────────────────────────
      sheet.setFrozenRows(1);

      // ── 2. Header row styling ─────────────────────────────────────────
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange
        .setBackground(CLR.headerBg)
        .setFontColor(CLR.headerFg)
        .setFontWeight('bold')
        .setFontSize(10)
        .setFontFamily('Google Sans, Arial, sans-serif')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(false);

      // Row height for header
      sheet.setRowHeight(1, 36);

      // ── 3. Alternating row banding ────────────────────────────────────
      if (lastRow > 1) {
        // Remove existing bandings first
        const bandings = sheet.getBandings();
        bandings.forEach(function(b) { b.remove(); });

        sheet.getRange(2, 1, lastRow - 1, lastCol)
          .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
          .setFirstRowColor(CLR.bandLight)
          .setSecondRowColor(CLR.bandDark)
          .setHeaderRowColor(null)
          .setFooterRowColor(null);
      }

      // ── 4. Data cell styling ──────────────────────────────────────────
      if (lastRow > 1) {
        const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
        dataRange
          .setFontSize(9)
          .setFontFamily('Google Sans, Arial, sans-serif')
          .setVerticalAlignment('middle')
          .setWrap(false);
        sheet.setRowHeightsForced(2, lastRow - 1, 28);
      }

      // ── 5. Borders on full table ──────────────────────────────────────
      const fullTable = sheet.getRange(1, 1, lastRow, lastCol);
      fullTable.setBorder(
        true, true, true, true, true, true,
        CLR.headerBorder,
        SpreadsheetApp.BorderStyle.SOLID
      );
      // Thicker outer border
      fullTable.setBorder(
        true, true, true, true, false, false,
        CLR.headerBg,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );

      // ── 6. Status column color-coding ─────────────────────────────────
      const meta = SHEET_META[name];
      if (meta && meta.statusCol && lastRow > 1) {
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const statusIdx = headers.findIndex(function(h) {
          return String(h).trim().toLowerCase() === meta.statusCol.toLowerCase();
        });
        if (statusIdx >= 0) {
          for (var r = 2; r <= lastRow; r++) {
            const cell = sheet.getRange(r, statusIdx + 1);
            const val  = String(cell.getValue()).trim().toLowerCase();
            var bg = '#FFFFFF';
            var fg = '#212529';
            if      (val === 'active' || val === 'paid' || val === 'approved' || val === 'sent')  { bg = CLR.statusActive; fg = '#155724'; }
            else if (val === 'inactive' || val === 'overdue' || val === 'rejected' || val === 'failed') { bg = CLR.statusInact; fg = '#721C24'; }
            else if (val === 'pending' || val === 'in progress' || val === 'renewal due')  { bg = CLR.statusPend; fg = '#856404'; }
            else if (val === 'sent' || val === 'issued')   { bg = CLR.statusSent; fg = '#0C5460'; }
            cell.setBackground(bg).setFontColor(fg).setFontWeight('bold').setHorizontalAlignment('center');
          }
        }
      }

      // ── 7. Auto-resize columns ────────────────────────────────────────
      for (var c = 1; c <= lastCol; c++) {
        sheet.autoResizeColumn(c);
        // Cap width for readability
        const w = sheet.getColumnWidth(c);
        if (w > 260) sheet.setColumnWidth(c, 260);
        if (w < 80)  sheet.setColumnWidth(c, 80);
      }

      // ── 8. Sheet tab color and emoji rename ───────────────────────────
      if (meta) {
        sheet.setTabColor(meta.tab);
        // Only rename if the sheet name doesn't already have an emoji
        const currentName = sheet.getName();
        if (!currentName.includes(meta.emoji)) {
          try { sheet.setName(meta.emoji + ' ' + currentName); } catch(e) { /* ignore duplicate name errors */ }
        }
      }

      results.push(name + ': formatted (' + lastRow + ' rows × ' + lastCol + ' cols)');
    });

    // ── 9. Add a "Summary Dashboard" sheet if not present ──────────────
    addSummaryDashboardSheet_(ss, CLR);

    SpreadsheetApp.flush();
    return createResponse(true, 'All sheets formatted successfully', { sheets: results });
  } catch (err) {
    return createResponse(false, 'formatAllSheets failed: ' + err.toString());
  }
}

/**
 * Adds (or refreshes) a "📊 Dashboard" sheet at the front of the workbook
 * with a KPI summary pulled live from the data sheets.
 */
function addSummaryDashboardSheet_(ss, CLR) {
  try {
    const DASH_NAME = '📊 Dashboard';
    let dash = ss.getSheetByName(DASH_NAME);
    if (!dash) {
      dash = ss.insertSheet(DASH_NAME, 0);
    } else {
      dash.clearContents();
      dash.clearFormats();
      ss.setActiveSheet(dash);
      ss.moveActiveSheet(0);
    }

    dash.setTabColor('#14213D');

    // ── Title block ───────────────────────────────────────────────────
    dash.getRange('A1:H1').merge()
      .setValue('🏢 MajengoSmart Pro — Property Management Dashboard')
      .setBackground('#14213D')
      .setFontColor('#F0C060')
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Google Sans, Arial, sans-serif')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    dash.setRowHeight(1, 52);

    dash.getRange('A2:H2').merge()
      .setValue('Auto-generated summary — Last updated: ' + Utilities.formatDate(new Date(), 'Africa/Nairobi', 'dd MMM yyyy HH:mm'))
      .setBackground('#1E3464')
      .setFontColor('#AAAAAA')
      .setFontSize(9)
      .setHorizontalAlignment('center');
    dash.setRowHeight(2, 24);

    // ── KPI section header ────────────────────────────────────────────
    dash.getRange('A4:H4').merge()
      .setValue('📈  KEY PERFORMANCE INDICATORS')
      .setBackground('#1A6B3A')
      .setFontColor('#FFFFFF')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    dash.setRowHeight(4, 32);

    // Pull live counts from sheets
    const tenantSheet  = ss.getSheetByName('Tenants')        || ss.getSheetByName('👥 Tenants');
    const paySheet     = ss.getSheetByName('RentPayments')   || ss.getSheetByName('💰 RentPayments');
    const invoiceSheet = ss.getSheetByName('Invoices')       || ss.getSheetByName('🧾 Invoices');
    const repairSheet  = ss.getSheetByName('Repairs')        || ss.getSheetByName('🔧 Repairs');

    function countRows(sh) { return sh && sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0; }
    function sumCol(sh, colName) {
      if (!sh || sh.getLastRow() < 2) return 0;
      const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = h.findIndex(function(c){ return String(c).trim().toLowerCase() === colName.toLowerCase(); });
      if (idx < 0) return 0;
      const vals = sh.getRange(2, idx+1, sh.getLastRow()-1, 1).getValues();
      return vals.reduce(function(s, r){ return s + (Number(r[0]) || 0); }, 0);
    }
    function countByStatus(sh, statusVal) {
      if (!sh || sh.getLastRow() < 2) return 0;
      const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      const idx = h.findIndex(function(c){ return String(c).toLowerCase() === 'status'; });
      if (idx < 0) return 0;
      const vals = sh.getRange(2, idx+1, sh.getLastRow()-1, 1).getValues();
      return vals.filter(function(r){ return String(r[0]).toLowerCase() === statusVal.toLowerCase(); }).length;
    }

    const totalTenants  = countRows(tenantSheet);
    const activeTenants = countByStatus(tenantSheet, 'Active');
    const totalRevenue  = sumCol(paySheet, 'Amount');
    const paidPayments  = countByStatus(paySheet, 'Paid');
    const pendingInv    = countByStatus(invoiceSheet, 'Pending') + countByStatus(invoiceSheet, 'Unpaid');
    const openRepairs   = countByStatus(repairSheet, 'Pending') + countByStatus(repairSheet, 'In Progress');

    // KPI cards in two rows of 4
    const kpis = [
      { label: '👥 Total Tenants',      value: totalTenants,           bg: '#14213D', fg: '#F0C060' },
      { label: '✅ Active Tenants',     value: activeTenants,          bg: '#1A6B3A', fg: '#FFFFFF' },
      { label: '💰 Total Revenue (KES)',value: 'KES ' + totalRevenue.toLocaleString(), bg: '#C8922A', fg: '#FFFFFF' },
      { label: '✔️ Paid Payments',      value: paidPayments,           bg: '#2D9E57', fg: '#FFFFFF' },
      { label: '🧾 Pending Invoices',   value: pendingInv,             bg: '#DC2626', fg: '#FFFFFF' },
      { label: '🔧 Open Repairs',       value: openRepairs,            bg: '#E65C00', fg: '#FFFFFF' },
      { label: '📋 All Invoices',       value: countRows(invoiceSheet),bg: '#0EA5E9', fg: '#FFFFFF' },
      { label: '🗂 All Repairs',        value: countRows(repairSheet), bg: '#7C3AED', fg: '#FFFFFF' },
    ];

    const kpiCols = ['A','B','C','D','E','F','G','H'];
    [5, 7].forEach(function(row, rowIdx) {
      kpis.slice(rowIdx*4, rowIdx*4+4).forEach(function(kpi, ci) {
        const col = kpiCols[ci];
        // Label
        dash.getRange(col + row).setValue(kpi.label)
          .setBackground(kpi.bg).setFontColor(kpi.fg)
          .setFontSize(8).setFontWeight('bold')
          .setHorizontalAlignment('center').setVerticalAlignment('bottom')
          .setWrap(false);
        dash.setRowHeight(row, 26);
        // Value
        dash.getRange(col + (row+1)).setValue(kpi.value)
          .setBackground(kpi.bg).setFontColor(kpi.fg)
          .setFontSize(18).setFontWeight('bold')
          .setFontFamily('Google Sans, Arial, sans-serif')
          .setHorizontalAlignment('center').setVerticalAlignment('top');
        dash.setRowHeight(row+1, 42);
      });
    });

    // ── Sheet index table ──────────────────────────────────────────────
    const idxRow = 10;
    dash.getRange('A' + idxRow + ':H' + idxRow).merge()
      .setValue('📂  SHEET INDEX')
      .setBackground('#1A6B3A').setFontColor('#FFFFFF')
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    dash.setRowHeight(idxRow, 32);

    dash.getRange('A' + (idxRow+1) + ':D' + (idxRow+1))
      .setValues([['Sheet Name', 'Rows of Data', 'Last Modified', 'Status']])
      .setBackground('#14213D').setFontColor('#F0C060')
      .setFontWeight('bold').setHorizontalAlignment('center');
    dash.setRowHeight(idxRow+1, 28);

    const dataSheets = ['Tenants','RentPayments','Invoices','Repairs','MaintenanceApprovals','TenantStatements','Aging_Tenants','Aging_Suppliers','Teams','Users'];
    dataSheets.forEach(function(sName, i) {
      var sh = ss.getSheetByName(sName);
      // Also try with emoji prefix
      if (!sh) sh = ss.getSheets().find(function(s){ return s.getName().includes(sName); });
      const rows = sh ? (sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0) : 'N/A';
      const statusTxt = sh ? '✅ Exists' : '⚠️ Missing';
      const bg = (i % 2 === 0) ? '#FFFFFF' : '#EBF4EF';
      const r = idxRow + 2 + i;
      dash.getRange(r, 1).setValue(sName);
      dash.getRange(r, 2).setValue(rows).setHorizontalAlignment('center');
      dash.getRange(r, 3).setValue(sh ? Utilities.formatDate(new Date(), 'Africa/Nairobi', 'dd MMM yyyy') : '-').setHorizontalAlignment('center');
      dash.getRange(r, 4).setValue(statusTxt).setHorizontalAlignment('center')
        .setFontColor(sh ? '#155724' : '#721C24').setFontWeight('bold');
      dash.getRange(r, 1, 1, 4).setBackground(bg).setFontSize(9);
      dash.setRowHeight(r, 24);
    });

    // Border entire index table
    const indexRange = dash.getRange(idxRow+1, 1, dataSheets.length+1, 4);
    indexRange.setBorder(true,true,true,true,true,true,'#14213D', SpreadsheetApp.BorderStyle.SOLID);

    // Auto-resize all columns
    for (var c = 1; c <= 8; c++) {
      dash.autoResizeColumn(c);
      if (dash.getColumnWidth(c) < 90) dash.setColumnWidth(c, 90);
    }

    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log('addSummaryDashboardSheet_ failed: ' + e.toString());
  }
}
