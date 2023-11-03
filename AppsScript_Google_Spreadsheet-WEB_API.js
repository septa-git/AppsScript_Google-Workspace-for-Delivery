//https://script.google.com/macros/s/AKfycbyPb7RJ2ZamRRBCeJMz7jR0mQDcHXkrfFz7DYi89E7Np8GOhPdNxaSoeLb37kiTYlzK/exec?sheet=INVOICES
// Fungsi untuk menangani permintaan GET ke endpoint API
function doGet(req) {
  // Membuat objek SpreadsheetApp untuk mengakses spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Membuat objek Sheet untuk mengakses lembar kerja "RIDERS" dan "INVOICES"
  var sheetRiders = ss.getSheetByName("RIDERS");
  var sheetInvoices = ss.getSheetByName("INVOICES");
  
  // Membuat objek ContentService untuk membuat output API
  var content = ContentService.createTextOutput();
  
  // Membuat objek JSON untuk menyimpan data dari lembar kerja
  var json = {
    data: []
  };
  
  // Mendapatkan parameter sheet dari permintaan
  var sheet = req.parameter.sheet;
  
  // Menentukan lembar kerja yang akan diproses berdasarkan parameter sheet
  var sheetObject;
  if (sheet == "RIDERS") {
    sheetObject = sheetRiders;
  } else if (sheet == "INVOICES") {
    sheetObject = sheetInvoices;
  } else {
    return response().json({ status: false, message: "Invalid sheet parameter" });
  }
  
  // Membuat variabel untuk menyimpan jumlah baris dan kolom pada lembar kerja
  var lastRow = sheetObject.getLastRow();
  var lastColumn = sheetObject.getLastColumn();
  
  // Menggunakan loop untuk membaca data dari setiap baris dan kolom pada lembar kerja
  for (var i = 2; i <= lastRow; i++) {
    // Membuat objek temp untuk menyimpan data dari satu baris
    var temp = {};
    // Menggunakan loop untuk membaca data dari setiap kolom pada baris tersebut
    for (var j = 1; j <= lastColumn; j++) {
      // Mendapatkan nama kolom dari baris pertama
      var key = sheetObject.getRange(1, j).getValue();
      // Mendapatkan nilai dari kolom tersebut pada baris tersebut
      var value = sheetObject.getRange(i, j).getValue();
      // Menyimpan pasangan nama dan nilai kolom ke objek temp
      temp[key] = value;
    }
    // Menambahkan objek temp ke array data pada objek json
    json.data.push(temp);
  }
  
  // Mengatur output API sebagai JSON dengan format yang indah (pretty)
  content.setMimeType(ContentService.MimeType.JSON);
  content.setContent(JSON.stringify(json, null, 2));
  
  // Mengembalikan output API sebagai respons
  return content;
}

// Fungsi bantuan untuk membuat respons JSON
function response() {
  return {
    json: function(data) {
      var result = ContentService.createTextOutput();
      result.setMimeType(ContentService.MimeType.JSON);
      result.setContent(JSON.stringify(data));
      return result;
    }
  };
}
