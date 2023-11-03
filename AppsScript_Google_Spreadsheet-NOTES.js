function onEdit(e) {
  var ss = e.source;
  var targetSheet = ss.getSheetByName('NOTES');
  
  // Cek apakah perubahan terjadi di sel M2 atau N2
  if (e.range.getSheet().getName() === 'NOTES' && (e.range.getA1Notation() === 'E5')) {
    if (e.value === true || e.value === "TRUE") {
      // Jeda eksekusi program selama 1 detik untuk menghindari masalah nonaktifkan checkbox segera
      Utilities.sleep(1000);
      NOTESDisplayAndFindValue();
      targetSheet.getRange('E5').setValue(false);
      // targetSheet.getRange('N2').setValue(false);
    }
  }
}


function NOTESDisplayAndFindValue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('GET_INVOICES');
  var targetSheet = ss.getSheetByName('NOTES');
  var testSheet = ss.getSheetByName('DUPLICATES');
  var riderSheet = ss.getSheetByName('COURIER');
  
  var filterValues = ["JK1 - JAKARTA TIMUR", "JK2 - JAKARTA BARAT", "DPK - DEPOK", "BDG - BANDUNG", "BTN - BANTEN/TANGERANG"];
  
  var range = sourceSheet.getRange('A:D');
  var values = range.getValues();
  
  targetSheet.getRange('F2:F').clearContent();
  targetSheet.getRange('AX2:AX').clearContent();
  
  var rowIndex = 2; // Mulai dari baris A2
  var uniqueInvoices = [];
  var duplicateInvoices = [];
  
  for (var i = 0; i < values.length; i++) {
    if (filterValues.includes(values[i][0])) {
      if (values[i][3]) {
        var invoices = values[i][3].split(',');
        for (var j = 0; j < invoices.length; j++) {
          var invoice = invoices[j].trim();
          if (!uniqueInvoices.includes(invoice) && !duplicateInvoices.includes(invoice)) { // Menggunakan kondisi if baru
            uniqueInvoices.push(invoice);
            targetSheet.getRange(rowIndex++, 6).setValue(invoice); // Mengubah indeks kolom menjadi 1 untuk kolom A
          } else if (!duplicateInvoices.includes(invoice)) {
            duplicateInvoices.push(invoice);
          }
        }
      }
    }
  }

  // Cek duplikat
  if (duplicateInvoices.length > 0) {
    targetSheet.getRange("E1").setValue("duplicate");
    testSheet.getRange('A2:A').clearContent();
    for (var i = 0; i < duplicateInvoices.length; i++) {
      testSheet.getRange(i + 2, 1).setValue(duplicateInvoices[i]);
      for (var j = 0; j < values.length; j++) {
        if (values[j][3] && values[j][3].includes(duplicateInvoices[i])) {
          testSheet.getRange(i + 2, 2).setValue(values[j][1]); // Menulis kolom B dari lembar kerja GET_INVOICES
          break;
        }
      }
    }
  } else {
    targetSheet.getRange("E5").clearContent();
    testSheet.getRange('A2:A').clearContent();
    testSheet.getRange('B2:B').clearContent();
  }

  
   // Ambil semua nilai sekaligus
   sourceValues = targetSheet.getRange("F2:F" + targetSheet.getLastRow()).getValues(); // Menambahkan tanda kutip di sekitar rentang sel
   
   // Ambil semua nilai sekaligus
   targetValues = sourceSheet.getRange("D2:D" + sourceSheet.getLastRow()).getValues(); // Menambahkan tanda kutip di sekitar rentang sel
   
   targetColumn3Values = sourceSheet.getRange("B2:B" + sourceSheet.getLastRow()).getValues(); // Menambahkan tanda kutip di sekitar rentang sel
   
   riderValues = riderSheet.getRange("C2:D" + riderSheet.getLastRow()).getValues(); // Menambahkan tanda kutip di sekitar rentang sel



   
   output = [];
   
   for (var i = 0; i < sourceValues.length; i++) {
     // Jika baris kosong, lewati
     if (!sourceValues[i][0]) {
       output.push([""]);
       continue;
     }
     
     found = false;
     for (var j = 0; j < targetValues.length; j++) {
       invoices = targetValues[j][0].split(', ');
       for (var k = 0; k < invoices.length; k++) {
         if (invoices[k] == sourceValues[i][0]) {
           // Cari nilai di lembar kerja RIDER
           for (var l = 0; l < riderValues.length; l++) {
             if (riderValues[l][0] == targetColumn3Values[j][0]) {
               output.push([riderValues[l][1]]);
               found = true;
               break;
             }
           }
         }
       }
       if (found) {
         break;
       }
     }
     if (!found) {
       output.push([""]);
     }
   }
   
    // Tulis semua output sekaligus
    targetSheet.getRange(2,52,output.length,1).setValues(output); // Mengubah kolom output menjadi AX (kolom ke-50)
}





        //POST MILEAPP
/*
function NOTESPostRequest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('NOTES');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var invoiceSheet = ss.getSheetByName('NOTE');
  var lastRow = invoiceSheet.getLastRow();
  var invoiceIndex = lastRow + 1;
  var existingInvoices = invoiceSheet.getRange('A2:A' + lastRow).getValues();

  var tasks = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var startTime = formatDateToISO(row[50]);
    var endTime = formatDateToISO(row[51]);
    var slotStart = formatDateToDateTime(row[28]);
    var slotEnd = formatDateToDateTime(row[29]);
    var consigneeSlotStart = formatDateToDateTime(row[41]);
    var consigneeSlotEnd = formatDateToDateTime(row[42]);
    var hubId = getHubId(row[20]);

    var task = {
      "flow": "Create Order Blitz",
      "hubId": hubId,
      "orderNumber": row[5],
      "customerName": row[6],
      "totalPackage": row[7],
      "volume": row[8],
      "volumeUnit": row[9],
      "width": row[10],
      "height": row[11],
      "length": row[12],
      "weight": row[13],
      "serviceType": row[14],
      "serviceRideType": row[15],
      "servicePromoCode": row[16],
      "servicePaymentType": row[17],
      "serviceCodAmount": row[18],
      "serviceTotalAmount": row[19],
      "senderName": row[20],
      "senderPhone": row[21],
      "senderDistrict": row[22],
      "senderCity": row[23],
      "senderProvince": row[24],
      "senderPostalcode": String(row[25]),
      "senderCountry": row[26],
      "senderLonglat": row[27],
      "senderSlotStart": slotStart,
      "senderSlotEnd": slotEnd,
      "senderEmail": row[30],
      "senderAddress": row[31],
      "senderNote": row[32],
      "consigneeName": row[33],
      "consigneePhone": row[34],
      "consigneeEmail": row[35],
      "consigneeDistrict": row[36],
      "consigneeCity": row[37],
      "consigneeProvince": row[38],
      "consigneePostalcode": row[39],
      "consigneeLonglat": row[40],
      "consigneeSlotStart": consigneeSlotStart,
      "consigneeSlotEnd": consigneeSlotEnd,
      "consigneeAddress": row[43],
      "consigneeCountry": row[44],
      "consigneeNote": row[45],
      "description": row[46],
      "productId": String(row[47]),
      "productName": row[48],
      "assignee": [row[49]],
      "startTime": startTime,
      "endTime": endTime
    };
    tasks.push(task);

    if (!existingInvoices.includes(row[0])) {
      invoiceSheet.getRange(invoiceIndex++, 1).setValue(row[5]);
    }
  }

  var url = 'https://apiweb.mile.app/api/v3/tasks/bulk/';
  var options = {
    'method' : 'post',
    'headers': {
      'Accept': 'application/json',
      'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiI2MmNmZWY2NDQ4NGMxZTcwMTU0YjQ5YjIiLCJqdGkiOiI5ZTg5NjYyOWUwNTNkNjhkNjIwNjNlNTFmNjQ1Yjk1MjEyZGE0NjcwNGVlY2RiOGIxYzQxOGI1MGQwZWUzZjFiOTBhOTZlNDdlM2JlN2NjNyIsImlhdCI6MTY5Mzk2NjM0MS42MTA1OTcsIm5iZiI6MTY5Mzk2NjM0MS42MTA2MDEsImV4cCI6NDg0OTYzOTk0MS42MDc2MzUsInN1YiI6IjY0Yzk0NDAyOTM0NGI4NTRiYjA2NTA3NSIsInNjb3BlcyI6W119.L_ex02T4qdm1v35xtFNPMbjz8EqCYx4qxYf6B2Se5aS6YpQC17sbslHtBKnLWYtlQUNJc6467rAemVoqQrpZs8ey_vZiAwoFVdgUE311wiwlnET9P6W83foh9xh4DBZR0Vs20fWBtRWpRZ3gJ21e95Eayan3fzsa11r-Uws8haoo_6oUJ3IyhlFUoteghFTvq24QGw0FRu3l90G0PTdS9G3W4Nfc4oH1gZxQ138H5DVhVwCO7GfBYTGluIZQXA2HbOLRDtXUTRhHQUSKGM-jzEi5dZ0YLbYjV5rhaqUbYUopkLq_YxrEfKxCZpgTzwdslZ5vlGA_FymS2PMHUPM_sxD0yfdtaJY9eyq0FFO2IsIzrDARApGCT4jZ-pcbbv784Fy1v-ZJcxZKFFf920bg3ugaWT8S2VJAuRKQ7QIzw_i04MrgDfoK75mp2IWwTnIT5B0M1-RmfnOP-7fRwBtDzm1dazzgaEDwPyEABmPV99NUPp3THBIhSLGvoqtZ9j2_CpwCGw35QXEPnmfjty9oE5T2N095_tkspE1G6Dj1GkELYAU5RFqjK3LTvfnpR5gbvRMO67H3Yp7lbmaRZCMzu1FtAJfM-d1oEyR7J2EJhf1UrW7w34gOyCMJgUJwf9YDVxHNzPJS6m9tpN88i-C1IKFAgiqK_S9KfRQwcXG9wIU',
      'Cache-Control': 'no-cache,private',
      'Content-Type': 'application/json'
    },
    'payload' : JSON.stringify({tasks: tasks})
  };
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);


  if (response.getResponseCode() == 200) {
    // Menyalin dan menempelkan isi fungsi postRequest2 di sini
  }
}
*/










function NOTESPostRequest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('NOTES');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var invoiceSheet = ss.getSheetByName('NOTE');
  var lastRow = invoiceSheet.getLastRow();
  var invoiceIndex = lastRow + 1;
  var existingInvoices = invoiceSheet.getRange('A2:A' + lastRow).getValues();

  var tasks = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var startTime = formatDateToISO(row[52]);
    var endTime = formatDateToISO(row[53]);
    var slotStart = formatDateToDateTime(row[29]);
    var slotEnd = formatDateToDateTime(row[30]);
    var consigneeSlotStart = formatDateToDateTime(row[43]);
    var consigneeSlotEnd = formatDateToDateTime(row[44]);
    var hubId = getHubId(row[20]);

    var task = {
      "flow": "Create Order Blitz",
      "hubId": hubId,
      "totalPackage_unit": "Package",
      "weight_unit": "KG",
      "orderNumber": row[5],
      "customerName": row[6],
      "totalPackage": row[7],
      "volume": row[8],
      "volumeUnit": row[9],
      "width": row[10],
      "height": row[11],
      "length": row[12],
      "weight": row[13],
      "serviceType": row[14],
      "serviceRideType": row[15],
      "servicePromoCode": row[16],
      "servicePaymentType": row[17],
      "serviceCodAmount": row[18],
      "serviceTotalAmount": row[19],
      "senderName": row[20],
      "senderPhone": row[21],
      "senderDistrict": row[22],
      "senderCity": row[23],
      "senderProvince": row[24],
      "senderPostalcode": String(row[25]),
      "senderCountry": row[26],
      "senderLonglat": row[27],
      "senderLatong": row[28],



      "senderSlotStart": slotStart,
      "senderSlotEnd": slotEnd,
      "senderEmail": row[31],
      "senderAddress": row[32],
      "senderNote": row[33],
      "consigneeName": row[34],
      "consigneePhone": row[35],
      "consigneeEmail": row[36],
      "consigneeDistrict": row[37],
      "consigneeCity": row[38],
      "consigneeProvince": row[39],
      "consigneePostalcode": row[40],
      "consigneeLonglat": row[41],
      "consigneeLatlong": row[42],

      "consigneeSlotStart": consigneeSlotStart,
      "consigneeSlotEnd": consigneeSlotEnd,
      "consigneeAddress": row[45],
      "consigneeCountry": row[46],
      "consigneeNote": row[47],
      "description": row[48],
      "productId": String(row[49]),
      "productName": row[50],
      "assignee": [row[51]],
      "startTime": startTime,
      "endTime": endTime
    };
    tasks.push(task);

    if (!existingInvoices.includes(row[0])) {
      invoiceSheet.getRange(invoiceIndex++, 1).setValue(row[5]);
    }
  }

  var url = 'https://apiweb.mile.app/api/v3/tasks/bulk/';
  var options = {
    'method' : 'post',
    'headers': {
      'Accept': 'application/json',
      'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiI2MmNmZWY2NDQ4NGMxZTcwMTU0YjQ5YjIiLCJqdGkiOiI5ZTg5NjYyOWUwNTNkNjhkNjIwNjNlNTFmNjQ1Yjk1MjEyZGE0NjcwNGVlY2RiOGIxYzQxOGI1MGQwZWUzZjFiOTBhOTZlNDdlM2JlN2NjNyIsImlhdCI6MTY5Mzk2NjM0MS42MTA1OTcsIm5iZiI6MTY5Mzk2NjM0MS42MTA2MDEsImV4cCI6NDg0OTYzOTk0MS42MDc2MzUsInN1YiI6IjY0Yzk0NDAyOTM0NGI4NTRiYjA2NTA3NSIsInNjb3BlcyI6W119.L_ex02T4qdm1v35xtFNPMbjz8EqCYx4qxYf6B2Se5aS6YpQC17sbslHtBKnLWYtlQUNJc6467rAemVoqQrpZs8ey_vZiAwoFVdgUE311wiwlnET9P6W83foh9xh4DBZR0Vs20fWBtRWpRZ3gJ21e95Eayan3fzsa11r-Uws8haoo_6oUJ3IyhlFUoteghFTvq24QGw0FRu3l90G0PTdS9G3W4Nfc4oH1gZxQ138H5DVhVwCO7GfBYTGluIZQXA2HbOLRDtXUTRhHQUSKGM-jzEi5dZ0YLbYjV5rhaqUbYUopkLq_YxrEfKxCZpgTzwdslZ5vlGA_FymS2PMHUPM_sxD0yfdtaJY9eyq0FFO2IsIzrDARApGCT4jZ-pcbbv784Fy1v-ZJcxZKFFf920bg3ugaWT8S2VJAuRKQ7QIzw_i04MrgDfoK75mp2IWwTnIT5B0M1-RmfnOP-7fRwBtDzm1dazzgaEDwPyEABmPV99NUPp3THBIhSLGvoqtZ9j2_CpwCGw35QXEPnmfjty9oE5T2N095_tkspE1G6Dj1GkELYAU5RFqjK3LTvfnpR5gbvRMO67H3Yp7lbmaRZCMzu1FtAJfM-d1oEyR7J2EJhf1UrW7w34gOyCMJgUJwf9YDVxHNzPJS6m9tpN88i-C1IKFAgiqK_S9KfRQwcXG9wIU',
      'Cache-Control': 'no-cache,private',
      'Content-Type': 'application/json'
    },
    'payload' : JSON.stringify({tasks: tasks})
  };
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);


  if (response.getResponseCode() == 200) {
    // Menyalin dan menempelkan isi fungsi postRequest2 di sini
  }
}







// Fungsi untuk mengubah format tanggal ke ISO
function formatDateToISO(date) {
  return Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

// Fungsi untuk mengubah format tanggal dan waktu
function formatDateToDateTime(dateTime) {
  return Utilities.formatDate(new Date(dateTime), "GMT", "yyyy-MM-dd HH:mm:ss");
}

// Fungsi untuk mendapatkan Hub ID berdasarkan nama
function getHubId(name) {
  var hubIds = {
    "PT MERAPI UTAMA PHARMA - JK1": "645b1d984c84e849033139f4",
    "PT MERAPI UTAMA PHARMA - JK2": "645b23a4a3e74b10574074e3",
    "PT MERAPI UTAMA PHARMA - DPK": "645b231d211f32fa0802e354",
    "PT MERAPI UTAMA PHARMA - BDG": "645b228073bcfc3f463541e2",
    "PT MERAPI UTAMA PHARMA - BTN": "645b22bfe355136c1234ca47"
  };

  return hubIds[name] || "";
}