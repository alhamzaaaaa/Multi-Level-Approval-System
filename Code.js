const ss = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_REQUESTS = "Request";      // Form responses sheet
const SHEET_CONFIG = "config";         // Role → Email mapping
const SHEET_RULES = "Rules";           // Amount → Chain mapping


//Helper function to get emails based on roles

function getApprovalEmail(Roles) {

  const configData = ss.getSheetByName(SHEET_CONFIG).getDataRange().getValues();

  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === Roles) {

      return configData[i][1];

    }
  }

  return null; // Role not found
}

function getApprovalChain(amount) {

  const rulesData = ss.getSheetByName(SHEET_RULES).getDataRange().getValues();
  let requiredChain = [];

  for (let i = rulesData.length - 1; i >= 0; i--) {

    let minAmount = Number(rulesData[i][0]);
    let chainString = rulesData[i][1];

    if (amount >= minAmount) {

      requiredChain = chainString.split(",").map(item => item.trim());
      break;

    }
  }

  return requiredChain;



}

function onFormSubmit(e) {


  Logger.log("e.values: " + JSON.stringify(e.values));

  const sheet = ss.getSheetByName(SHEET_REQUESTS);

  const row = e.values;

  const employeeEmail = row[4];
  const amount = Number(row[2]);

  const chain = getApprovalChain(amount);

  if (chain.length === 0) {

    Logger.log("No Chain Found For Amount: " + amount);
    return;
  }

  const firstApprover = chain[0];

  const lastRow = sheet.getLastRow();

  sheet.getRange(lastRow, 6).setValue("Pending");
  sheet.getRange(lastRow, 7).setValue(firstApprover);
  sheet.getRange(lastRow, 8).setValue(JSON.stringify(chain));

  Logger.log("New Request From: " + employeeEmail);
  Logger.log("Amount: " + amount);
  Logger.log("Chain: " + JSON.stringify(chain));
  Logger.log("First Approver: " + firstApprover);

  const approverEmail = getApprovalEmail(firstApprover);
  sendApprovalEmail(approverEmail, lastRow);


}


function sendApprovalEmail(approverEmail, rowNumber) {

  const webAppUrl = "https://script.google.com/macros/s/AKfycbzUJk35Iw4Jpv2PVz8E2djjODNX99pZKTJPbd-9AokKERo7b2p5X7wnGXD5HNiPvf-e/exec";

  const approveUrl = `${webAppUrl}?action=approve&row=${rowNumber}`;
  const rejectUrl = `${webAppUrl}?action=approve&row=${rowNumber}`;


  const sheet = ss.getSheetByName(SHEET_REQUESTS);

  const rowData = sheet.getRange(rowNumber, 1, 1, 8).getValues()[0];

  const subject = `Approval Required:${rowData[1]} + " of " ${rowData[2]} `;

  const htmlBody = `
      <p>Hello,</p>
      <p>A new request awaits your approval:</p>
      <ul>
        <li><b>Employee:</b> ${rowData[4]}</li>
        <li><b>Type:</b> ${rowData[1]}</li>
        <li><b>Amount:</b> ₹${rowData[2]}</li>
        <li><b>Description:</b> ${rowData[3]}</li>
      </ul>
      <p>
        <a href="${approveUrl}" style="background:#28a745;color:#fff;padding:10px 20px;text-decoration:none;border-radius:5px;">✅ Approve</a>
        &nbsp;&nbsp;
        <a href="${rejectUrl}" style="background:#dc3545;color:#fff;padding:10px 20px;text-decoration:none;border-radius:5px;">❌ Reject</a>
      </p>
  `;

  GmailApp.sendEmail(approverEmail, subject, "Please enable HTML to view this email.", {
    htmlBody: htmlBody
  });
  Logger.log("Email sent to: " + approverEmail);


}


function doGet(e) {

  const action = e.parameter.action;
  const row = Number(e.parameter.row);
  
  Logger.log("Action: " + action + " | Row: " + row);


  return HtmlService.createHtmlOutput(
    "<h2> Decision Recorded. </h2><p> You can close this tab.</p>"
  );


}



















































