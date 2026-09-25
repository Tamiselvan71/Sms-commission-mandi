/* GOOGLE APPS SCRIPT — Code.gs */

const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === "load") {
    const data = readDatabase();
    return json({ok:true,data:data});
  }

  return json({
    ok:true,
    message:"SMS Commission Mandi Cloud API is running"
  });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");

    if (body.action === "save") {
      writeDatabase(body.data || {});
      return json({ok:true,message:"Saved"});
    }

    return json({ok:false,message:"Unknown action"});
  } catch(err) {
    return json({ok:false,error:String(err)});
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName("SMS_DATA");

  if (!sh) {
    sh = ss.insertSheet("SMS_DATA");
    sh.getRange(1,1,1,3).setValues([
      ["KEY","VALUE","UPDATED"]
    ]);
  }

  return sh;
}

function writeDatabase(data) {
  const sh = getSheet();

  const rows = [
    ["customers",JSON.stringify(data.customers || []),new Date()],
    ["bills",JSON.stringify(data.bills || []),new Date()],
    ["receipts",JSON.stringify(data.receipts || []),new Date()],
    ["purchases",JSON.stringify(data.purchases || []),new Date()],
    ["partyPayments",JSON.stringify(data.partyPayments || []),new Date()],
    ["billSettings",JSON.stringify(data.billSettings || {}),new Date()]
  ];

  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([["KEY","VALUE","UPDATED"]]);
  sh.getRange(2,1,rows.length,3).setValues(rows);
}

function readDatabase() {
  const sh = getSheet();
  const values = sh.getDataRange().getValues();

  const data = {
    customers:[],
    bills:[],
    receipts:[],
    purchases:[],
    partyPayments:[],
    billSettings:{
      header:"SMS COMMISSION MANDI",
      subHeader:"Customer Accounts & Billing",
      contact:"",
      footer:"Thank You"
    }
  };

  for (let i=1;i<values.length;i++) {
    const key = values[i][0];
    const value = values[i][1];

    if (!key || !value) continue;

    try {
      data[key] = JSON.parse(value);
    } catch(err) {}
  }

  return data;
}
