      //kasa1
function dataFTPHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('BACKUP_FTP_HISTORY');
  var targetSheet = ss.getSheetByName('FTP_HISTORY');
  
  var companies = ['PT MERAPI UTAMA PHARMA - JK1', 'PT MERAPI UTAMA PHARMA - JK2', 'PT MERAPI UTAMA PHARMA - DPK', 'PT MERAPI UTAMA PHARMA - BDG', 'PT MERAPI UTAMA PHARMA - BTN'];
  
  var data = sourceSheet.getDataRange().getValues();
  
  data.forEach(function(row) {
    if(companies.includes(row[6])) { // Kolom E adalah indeks ke-4
      var targetRow = Array(targetSheet.getLastColumn()).fill(""); // Membuat array dengan panjang yang sama dengan jumlah kolom di targetSheet, diisi dengan string kosong
      targetRow.splice(0, row.length, ...row); // Menambahkan baris data ke targetRow, dimulai dari indeks ke-2 (kolom C)
      targetSheet.appendRow(targetRow);
    }
  });
}





      //kasa2
function googleForm() {
  // Mendefinisikan spreadsheet yang aktif
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Mendefinisikan lembar kerja yang akan diolah
  var sheets = ["NDS", "SDS1", "SDS2"];
  // Mendefinisikan lembar kerja yang akan menampilkan hasil
  var outputSheet = spreadsheet.getSheetByName("DATA_API");
  // Mendapatkan jumlah baris yang sudah ada di lembar kerja output
  var lastRow = outputSheet.getLastRow();
  // Mendefinisikan value yang dicari
  var values = ["PT MERAPI UTAMA PHARMA - JK1", "PT MERAPI UTAMA PHARMA - JK2", "PT MERAPI UTAMA PHARMA - DPK", "PT MERAPI UTAMA PHARMA - BDG", "PT MERAPI UTAMA PHARMA - BTN"];
  // Melakukan iterasi untuk setiap lembar kerja yang akan diolah
  for (var i = 0; i < sheets.length; i++) {
    // Mendapatkan data dari lembar kerja yang sedang diiterasi
    var sheet = spreadsheet.getSheetByName(sheets[i]);
    var data = sheet.getDataRange().getValues();
    // Melakukan iterasi untuk setiap baris data
    for (var j = 0; j < data.length; j++) {
      // Mendapatkan value dari kolom E (index ke-4)
      var valueE = data[j][4];
      // Memeriksa apakah value termasuk dalam daftar value yang dicari
      if (values.includes(valueE)) {
        // Menambahkan value per baris ke lembar kerja output tanpa menimpa data sebelumnya
        outputSheet.getRange(lastRow + 1, 1, 1, data[j].length).setValues([data[j]]);
        // Menambahkan nilai lastRow untuk mengisi baris selanjutnya
        lastRow++;
        // Inisialisasi variabel untuk menyimpan value yang akan diberikan pada kolom X dan Y
        var valueX;
        var valueY;
        // Logika untuk menentukan value berdasarkan value pada kolom E
        if (valueE == "PT MERAPI UTAMA PHARMA - JK1") {
          // Jika value pada kolom E adalah "PT MERAPI UTAMA PHARMA - JK1", maka berikan value "13930" pada kolom X dan value "-6.2687694,106.8028143" pada kolom Y
          valueX = "13930";
          valueY = "-6.2687694,106.8028143";
        } else if (valueE == "PT MERAPI UTAMA PHARMA - JK2") {
          // Jika value pada kolom E adalah "PT MERAPI UTAMA PHARMA - JK2", maka berikan value "11850" pada kolom X dan value "-6.2666795,106.7972406" pada kolom Y
          valueX = "11850";
          valueY = "-6.2666795,106.7972406";
        } else if (valueE == "PT MERAPI UTAMA PHARMA - DPK") {
          // Jika value pada kolom E adalah "PT MERAPI UTAMA PHARMA - DPK", maka berikan value "16453" pada kolom X dan value "-6.2666795,106.7972406" pada kolom Y
          valueX = "16453";
          valueY = "-6.2666795,106.7972406";
        } else if (valueE == "PT MERAPI UTAMA PHARMA - BDG") {
          // Jika value pada kolom E adalah "PT MERAPI UTAMA PHARMA - BDG", maka berikan value "40295" pada kolom X dan value "-6.9468451,107.6476942" pada kolom Y
          valueX = "40295";
          valueY = "-6.9468451,107.6476942";
        } else if (valueE == "PT MERAPI UTAMA PHARMA - BTN") {
          // Jika value pada kolom E adalah "PT MERAPI UTAMA PHARMA - BTN", maka berikan value "15113" pada kolom X dan value "-6.1851673,106.6212918" pada kolom Y
          valueX = "15113";
          valueY = "-6.1851673,106.6212918";
        } else {
          // Jika value pada kolom E tidak sesuai dengan ketiga kondisi di atas, maka berikan value kosong pada kolom X dan Y
          valueX = "";
          valueY = "";
        }
        // Menyimpan value yang telah ditentukan pada array nilai kolom X dan Y
        outputSheet.getRange(lastRow, 24, 1, 1).setValue(valueX);
        outputSheet.getRange(lastRow, 25, 1, 1).setValue(valueY);
      }
    }
  }
}





      //kasa3
function waSendFTP() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName('FIND_INVOICE');
  
    var configs = [
    {chatId: "120363046004663946@g.us", columns: [['A', 'B'], ['E', 'F'], ['I', 'J']], text: "Jangan lupa untuk input GFORM, ya Pak  @6289517188882 @6289670230682 @6287794900091 @6281296893653 @6288293457214 \n\n", mentions: ["6289517188882@c.us", "6289670230682@c.us", "6287794900091@c.us", "6281296893653@c.us", "6288293457214@c.us"]},
    {chatId: "120363029500708427@g.us", columns: [['M', 'N'], ['Q', 'R'], ['U', 'V']], text: "Jangan lupa untuk input GFORM, ya Pak @6281212983998 @6283892580874 @6281952747810 @6281386813380 @6281295604504 @6281384482133 \n\n", mentions: ["6281212983998@c.us", "6283892580874@c.us", "6281952747810@c.us", "6281386813380@c.us", "6281295604504@c.us", "6281384482133@c.us"]},
    {chatId: "120363030339996986@g.us", columns: [['Y', 'Z'], ['AC', 'AD'], ['AG', 'AH']], text: "Jangan lupa untuk input GFORM, ya Pak @6281211718860 @6287884417356 @6281212090053 @628567344775 \n\n", mentions: ["6281211718860@c.us", "6287884417356@c.us", "6281212090053@c.us", "628567344775@c.us"]},
    {chatId: "120363081282848203@g.us", columns: [['AK', 'AL'], ['AO', 'AP'], ['AS', 'AT']], text: "Jangan lupa untuk input GFORM, ya Pak @6283142682373 @6282262062462 @6281287560747 @628814093715 @6281395098566 \n\n", mentions: ["6283142682373@c.us", "6282262062462@c.us", "6281287560747@c.us", "628814093715@c.us", "6281395098566@c.us"]},
    {chatId: "120363079304273658@g.us", columns: [['AW', 'AX'], ['BA', 'BB'], ['BE', 'BF']], text: "Jangan lupa untuk input GFORM, ya Pak @6281296442653 @6285692018738 @6281385925335 \n\n", mentions: ["6281296442653@c.us", "6285692018738@c.us", "6281385925335@c.us"]}
  ];
  
  for (var i = 0; i < configs.length; i++) {
    var config = configs[i];
    var text = '';
    for (var j = 0; j < config.columns.length; j++) {
      var range = sourceSheet.getRange(config.columns[j][0] + ':' + config.columns[j][1]);
      var values = range.getValues();
      if (!values[1][0] || !values[1][1]) {
        continue;
      }
      for (var k = 0; k < values.length; k++) {
        if (values[k][0] && values[k][1]) {
          text += values[k].join(', ') + '\n';
        }
      }
      text += "\n\n";
    }
    
    if (text.trim() !== '') {
      var url = 'https://dennis-poem-easy-outlet.trycloudflare.com/api/sendText';
      var payload = {
        "chatId": config.chatId,
        "mentions": config.mentions,
        "text": config.text + text,
        "session": "default"
      };
      
      var options = {
        'method' : 'post',
        'headers': {
          'accept': 'application/json',
          'Content-Type': 'application/json'
        },
        'payload' : JSON.stringify(payload)
      };
      
      UrlFetchApp.fetch(url, options);
    }
  }
}
