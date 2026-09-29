// URL ID Spreadsheet Anda
var sheetId = '1f6RC9Vhy1PXEeEXoN_udz50Dg3dXuZ-PmEZKsYbqgmY'; 

// Router API untuk menerima request dari GitHub Pages
function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;
    var data = contents.data;
    var result = {};

    if (action === 'cekLogin') {
      result = cekLogin(data.username, data.password);
    } else if (action === 'simpanDataSiswa') {
      result = simpanDataSiswa(data);
    } else if (action === 'getSemuaSiswa') {
      result = getSemuaSiswa();
    } else if (action === 'simpanSiswaBaru') {
      result = simpanSiswaBaru(data.nisn, data.nama, data.kelas);
    } else if (action === 'editDataSiswa') {
      result = editDataSiswa(data.nisnLama, data.nisnBaru, data.namaBaru, data.kelasBaru);
    } else if (action === 'hapusDataSiswa') {
      result = hapusDataSiswa(data.nisn);
    } else if (action === 'getRekapIbadah') {
      result = getRekapIbadah();
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'API Jurnal Ibadah Active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================= DATABASE FUNCTIONS =================
function cekLogin(username, password) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === username && data[i][3].toString() === password) {
      return {
        status: 'sukses',
        nama: data[i][1],
        kelas: data[i][2],
        role: data[i][4]
      };
    }
  }
  return { status: 'gagal' };
}

function simpanDataSiswa(data) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Data_Ibadah');
  var waktuSekarang = new Date();
  
  sheet.appendRow([
    waktuSekarang,
    data.nisn,
    data.nama,
    data.subuh,
    data.dhuhur,
    data.ashar,
    data.maghrib,
    data.isya,
    data.tahajjud,
    data.dhuha
  ]);
  
  return "Alhamdulillah, laporan ibadah berhasil disimpan!";
}

function getSemuaSiswa() {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === 'Siswa') {
      hasil.push({ nisn: data[i][0], nama: data[i][1], kelas: data[i][2] });
    }
  }
  return hasil;
}

function simpanSiswaBaru(nisn, nama, kelas) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Users');
  var passwordSiswa = nisn; 
  sheet.appendRow([nisn, nama, kelas, passwordSiswa, 'Siswa']);
  return "Siswa berhasil ditambahkan!";
}

function editDataSiswa(nisnLama, nisnBaru, namaBaru, kelasBaru) {
   var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Users');
   var data = sheet.getDataRange().getValues();
   for (var i = 1; i < data.length; i++) {
     if (data[i][0].toString() === nisnLama.toString() && data[i][4] === 'Siswa') {
       var row = i + 1;
       sheet.getRange(row, 1).setValue(nisnBaru);
       sheet.getRange(row, 2).setValue(namaBaru);
       sheet.getRange(row, 3).setValue(kelasBaru);
       sheet.getRange(row, 4).setValue(nisnBaru);
       return "Data siswa berhasil diupdate!";
     }
   }
   return "Data tidak ditemukan.";
}

function hapusDataSiswa(nisn) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nisn.toString() && data[i][4] === 'Siswa') {
      sheet.deleteRow(i + 1);
      return "Siswa berhasil dihapus!";
    }
  }
  return "Siswa tidak ditemukan.";
}

function getRekapIbadah() {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Data_Ibadah');
  var data = sheet.getDataRange().getDisplayValues(); 
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    hasil.push({
      waktu: data[i][0], nisn: data[i][1], nama: data[i][2],
      subuh: data[i][3], dhuhur: data[i][4], ashar: data[i][5], 
      maghrib: data[i][6], isya: data[i][7], tahajjud: data[i][8], dhuha: data[i][9]
    });
  }
  return hasil.reverse();
}