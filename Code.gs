/* ============================================================
   GANTI DENGAN ID GOOGLE SHEET ANDA
   Contoh: "1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890"
   ============================================================ */
const SHEET_ID = "17-u6pQcrjf7YX-eKP_CqaVakM0_yNQ3oMG_AZ0qR3To";
const SHEET_NAME = "DataPelayanan";

function doGet(e) {
  return jsonResponse({ success: true, data: getAllData() });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, data } = body;
    let msg = "";
    if (action === "create") { msg = createData(data); }
    else if (action === "update") { msg = updateData(data); }
    else if (action === "delete") { msg = deleteData(data.id); }
    else throw new Error("Action tidak valid");
    return jsonResponse({ success: true, message: msg });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["ID","No Pelayanan","NOP","Nama","Pemohon","PPAT","Jenis Pelayanan","Keterangan","Timestamp"]);
  }
  return sh;
}

function getAllData() {
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(r => ({
    id: r[0],
    noPelayanan: r[1],
    nop: r[2],
    nama: r[3],
    pemohon: r[4],
    ppat: r[5],
    jenisPelayanan: r[6],
    keterangan: r[7]
  }));
}

function createData(d) {
  const sh = getSheet();
  const id = "PL-" + new Date().getTime();
  sh.appendRow([id, d.noPelayanan, d.nop, d.nama, d.pemohon, d.ppat, d.jenisPelayanan, d.keterangan, new Date()]);
  return "✅ Data berhasil ditambahkan";
}

function updateData(d) {
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == d.id) {
      sh.getRange(i+1, 1, 1, 8).setValues([[d.id, d.noPelayanan, d.nop, d.nama, d.pemohon, d.ppat, d.jenisPelayanan, d.keterangan]]);
      return "✅ Data berhasil diperbarui";
    }
  }
  throw new Error("Data tidak ditemukan");
}

function deleteData(id) {
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == id) {
      sh.deleteRow(i+1);
      return "✅ Data berhasil dihapus";
    }
  }
  throw new Error("Data tidak ditemukan");
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}