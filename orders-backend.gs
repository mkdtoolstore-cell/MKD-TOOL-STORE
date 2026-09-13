/**
 * LaoBuild — Orders backend (Google Apps Script)
 * =================================================================
 * ວິທີຕິດຕັ້ງ:
 * 1. ສ້າງ Google Sheet ໃໝ່ (ຫ້ອງວ່າງໆ ກໍ່ໄດ້)
 * 2. ເມນູ Extensions → Apps Script
 * 3. ລຶບໂຄ້ດຕົວຢ່າງ ແລ້ວວາງໂຄ້ດທັງໝົດນີ້ລົງແທນ
 * 4. Deploy → New deployment → ເລືອກ "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    ແລ້ວກົດ Deploy (ຈະຕ້ອງ Authorize ສິດເຂົ້າເຖິງ Sheet/Drive ຂອງທ່ານ)
 * 5. ກັອບປີ້ລິ້ງ "Web app URL" (ລົງທ້າຍດ້ວຍ /exec)
 * 6. ນຳລິ້ງນັ້ນໄປວາງໃສ່ຕົວແປ ORDERS_API_URL ໃນທັງ index.html ແລະ admin.html
 *
 * Sheet "Orders" ຈະຖືກສ້າງອັດຕະໂນມັດພ້ອມຫົວຕາຕະລາງໃນຄັ້ງທຳອິດທີ່ມີການສັ່ງຊື້.
 * ຮູບສະລິບການໂອນເງິນຈະຖືກເກັບໄວ້ໃນໂຟນເດີ Google Drive ຊື່ "LaoBuild_Slips"
 * (ສ້າງອັດຕະໂນມັດ) ແລະ ເກັບແຕ່ລິ້ງຮູບໄວ້ໃນ Sheet (ບໍ່ໄດ້ເກັບຮູບຕົວເຕັມໃນ Sheet).
 * =================================================================
 */

const SHEET_NAME = "Orders";
const SLIP_FOLDER_NAME = "LaoBuild_Slips";
const HEADERS = ["code","createdAt","name","phone","addr","payMethod","items","total","status","slipImage"];

function doPost(e){
  try{
    const payload = JSON.parse(e.postData.contents);
    if(payload.action === "newOrder"){
      return jsonOut(handleNewOrder(payload.order));
    }
    if(payload.action === "updateStatus"){
      return jsonOut(handleUpdateStatus(payload.code, payload.status));
    }
    return jsonOut({ok:false, error:"unknown action"});
  }catch(err){
    return jsonOut({ok:false, error:String(err)});
  }
}

function doGet(e){
  try{
    const action = (e.parameter.action || "list");
    if(action === "track"){
      return jsonOut(handleTrack(e.parameter.code || ""));
    }
    return jsonOut(handleList());
  }catch(err){
    return jsonOut({ok:false, error:String(err)});
  }
}

/* ---------------- ACTIONS ---------------- */
function handleNewOrder(order){
  if(!order || !order.code) return {ok:false, error:"missing order"};
  const sheet = getSheet();
  let slipUrl = "";
  if(order.slipImage){
    slipUrl = saveSlipImage(order.slipImage, order.code);
  }
  sheet.appendRow([
    order.code,
    new Date(order.createdAt || Date.now()),
    order.name || "",
    order.phone || "",
    order.addr || "",
    order.payMethod || "",
    JSON.stringify(order.items || []),
    order.total || 0,
    order.status || "ລໍຖ້າກວດສອບ",
    slipUrl
  ]);
  return {ok:true, code: order.code};
}

function handleUpdateStatus(code, status){
  if(!code || !status) return {ok:false, error:"missing code/status"};
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(code)){
      sheet.getRange(i+1, HEADERS.indexOf("status")+1).setValue(status);
      return {ok:true};
    }
  }
  return {ok:false, error:"order not found"};
}

function handleList(){
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const orders = [];
  for(let i=1; i<data.length; i++){
    orders.push(rowToOrder(data[i]));
  }
  return {ok:true, orders};
}

function handleTrack(code){
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]).toUpperCase() === String(code).toUpperCase()){
      return {ok:true, order: rowToOrder(data[i])};
    }
  }
  return {ok:false, error:"not found"};
}

/* ---------------- HELPERS ---------------- */
function rowToOrder(row){
  let items = [];
  try{ items = JSON.parse(row[6] || "[]"); }catch(e){ items = []; }
  return {
    code: row[0],
    createdAt: new Date(row[1]).getTime(),
    name: row[2],
    phone: row[3],
    addr: row[4],
    payMethod: row[5],
    items: items,
    total: row[7],
    status: row[8],
    slipImage: row[9] || null
  };
}

function getSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function getSlipFolder(){
  const folders = DriveApp.getFoldersByName(SLIP_FOLDER_NAME);
  if(folders.hasNext()) return folders.next();
  return DriveApp.createFolder(SLIP_FOLDER_NAME);
}

function saveSlipImage(dataUrl, code){
  try{
    const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
    if(!match) return "";
    const mimeType = match[1];
    const base64 = match[2];
    const bytes = Utilities.base64Decode(base64);
    const ext = mimeType.split("/")[1] || "jpg";
    const blob = Utilities.newBlob(bytes, mimeType, "slip_" + code + "." + ext);
    const folder = getSlipFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/uc?id=" + file.getId();
  }catch(err){
    return "";
  }
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
