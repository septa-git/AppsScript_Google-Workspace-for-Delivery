function onSubmit(e) {
  // Mendapatkan form dan response
  var form = e.source;
  var response = e.response;
  
  // Mendapatkan semua item Multiple Choice
  var mcItems = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
  // Membuat array baru untuk menyimpan nilai dari semua jawaban Multiple Choice
  var mcValues = [];
  // Looping untuk setiap item Multiple Choice
  for (var i = 0; i < mcItems.length; i++) {
    // Mendapatkan item saat ini
    var mcItem = mcItems[i];
    // Mendapatkan jawaban dari item saat ini
    var mcAnswer = response.getResponseForItem(mcItem);
    // Jika jawaban tidak kosong
    if (mcAnswer) {
      // Mendapatkan nilai dari jawaban
      var mcValue = mcAnswer.getResponse();
      // Menambahkan nilai tersebut ke array mcValues
      mcValues.push(mcValue);
    }
  }
  
  // Mendapatkan item Checkbox
  var items = form.getItems(FormApp.ItemType.CHECKBOX);
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var answer = response.getResponseForItem(item);
    if (answer) {
      var values = answer.getResponse();
      // Mengambil data dari API menggunakan fungsi getDataFromAPI
      var data = getDataFromAPI("https://septa-git.github.io/db1/data.json");
      // var data = getDataFromAPI("https://script.googleusercontent.com/macros/echo?user_content_key=i6oqyE_atxc8mrn_t7ialmpGDaGXH_5JtEg8ZnJXI-c0ncCQMf2a9GM5qix9eu513I6D1OXXZRelRCEucWa3Opm7Yp5GdDPrm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnI_-srDRVqQrHiJ5rhsJh_Zomz_MlDIG9FQR_FomOMxRIyZH8-4wnyJb8fDNOVcEZ6GYcf4N3x-E2Lv47ns6m34j3NmcT0s4L58C3JM0WqjHCMY51C7ETn0&lib=MZ5jSiAJVPsiSLHrqedSLQRSYkE5FYcYl");
      for (var j = 0; j < values.length; j++) {
        var value = values[j];
        for (var k = 0; k < data.length; k++) {
          var obj = data[k];
          if (value == obj.sell_id) {
            // Memanggil fungsi PostRequest dengan parameter yang sesuai
            PostRequest(value, obj.packages, obj.co_name, obj.co_mobile, obj.co_count, obj.co_city, obj.co_prov, obj.senderPostalcode, obj.senderLonglat, obj.senderLatong, obj.co_add, obj.coe_name, obj.coe_mobile, obj.coe_count, obj.coe_city, obj.coe_prov, obj.coe_add, obj.so_num, obj.artc, mcValues); // Menambahkan parameter mcValues untuk menyimpan nilai dari semua jawaban Multiple Choice
            break;
          }
        }
      }
    }
  }
}

function getDataFromAPI(url) {
  // Mengirimkan permintaan ke URL yang diberikan dan mengembalikan respons dalam bentuk JSON
  var response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText()).data;
}

function PostRequest(faktur, totalPackage, cabang, senderPhone, senderDistrict, senderCity, senderProvince, senderPostalcode, senderLonglat, senderLatong, senderAddress, consigneeName, consigneePhone, consigneeDistrict, consigneeCity, consigneeProvince, consigneeAddress, productId, productName, mcValues) {
    var tasks = [];
    // Mendapatkan data dari API lain
    var data2 = getDataFromAPI("https://script.googleusercontent.com/macros/echo?user_content_key=5F6fSRnPOkzmiRj1WNtZg4d3UULXqRVQDGZgFQQDinawXSRacQQ4M1gNSZqAETmxibUe2SRarTyLiLP4T3vT5Fa6MY6lV_mWm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnIr_bsyuptVh-ooOw8YinfFO_2-Rv3ubXF_XZpp84xPTwK2GCh_UezGyEZ3usp8X1OjhObRyO0rjPs3bl44sgczUxYZG7R1Y7KHl1-rOccv9Z-r_BKdwmyY&lib=MrH4ffIkpvtg3OvNMCDW4bcoK9_9-SB1U");
    // Membuat variabel baru untuk menyimpan nilai email yang cocok dengan nilai rider
    var email = "";
    // Looping untuk setiap objek dalam data2
    for (var j = 0; j < data2.length; j++) {
      // Mendapatkan objek saat ini
      var obj2 = data2[j];
      // Looping untuk setiap nilai dalam array mcValues
      for (var i = 0; i < mcValues.length; i++) {
        // Mendapatkan nilai saat ini
        var mcValue = mcValues[i];
        // Jika nilai mcValue sama dengan nilai rider dari objek saat ini
        if (mcValue == obj2.rider) {
          // Mendapatkan nilai email dari objek saat ini
          email = obj2.email;
          break; // Keluar dari loop
        }
      }
    }
    // Membuat variabel baru untuk menyimpan nilai slot yang cocok dengan nilai NDS, SDS1, atau SDS2
    var slot = "";
    // Looping untuk setiap nilai dalam array mcValues
    for (var i = 0; i < mcValues.length; i++) {
      // Mendapatkan nilai saat ini
      var mcValue = mcValues[i];
      // Jika nilai mcValue sama dengan "NDS", "SDS1", atau "SDS2"
      if (mcValue == "NDS" || mcValue == "SDS1" || mcValue == "SDS2") {
        // Mendapatkan nilai slot dari nilai mcValue
        slot = mcValue;
        break; // Keluar dari loop
      }
    }







var senderSlotStart = "";
var senderSlotEnd = "";

// Menggunakan switch case untuk menentukan nilai senderSlotStart dan senderSlotEnd berdasarkan nilai slot
switch (slot) {
  case "NDS":
    // Jika slot adalah NDS, maka senderSlotStart adalah jam 17:00 PM dan senderSlotEnd adalah jam 11:00 AM
    senderSlotStart = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'17:00:00XXX');
    senderSlotEnd = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'11:00:00XXX');
    break;
  case "SDS1":
    // Jika slot adalah SDS1, maka senderSlotStart adalah jam 11:00 AM dan senderSlotEnd adalah jam 14:30 PM
    senderSlotStart = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'11:00:00XXX');
    senderSlotEnd = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'14:30:00XXX');
    break;
  case "SDS2":
    // Jika slot adalah SDS2, maka senderSlotStart adalah jam 14:30 PM dan senderSlotEnd adalah jam 17:00 PM
    senderSlotStart = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'14:30:00XXX');
    senderSlotEnd = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'17:00:00XXX');
    break;
  default:
    // Jika slot tidak ada dalam pilihan, maka senderSlotStart dan senderSlotEnd tetap kosong
    senderSlotStart = "";
    senderSlotEnd = "";
}




    
    // Membuat objek task dengan properti yang sesuai
    var task = {
      "flow": "Create Order Blitz",
      "hubId": getHubId(cabang),
      "totalPackage_unit": "Package",
      "weight_unit": "KG",
      "orderNumber": faktur,
      "customerName": "Merapi Utama Pharma",
      "totalPackage": totalPackage,
      "volume": 1,
      "volumeUnit": "cm3",
      "width": 1,
      "height": 1,
      "length": 1,
      "weight": 1,
      "serviceType": slot,
      "serviceRideType": "BIKE",
      "servicePromoCode": null,
      "servicePaymentType": "NON COD",
      "serviceCodAmount": 0,
      "serviceTotalAmount": 1,
      "senderName": cabang,
      "senderPhone": senderPhone,
      "senderDistrict": senderDistrict,
      "senderCity": senderCity,
      "senderProvince": senderProvince,
      "senderPostalcode": String(senderPostalcode),
      "senderCountry": "Indoneisa",
      "senderLonglat": String(senderLonglat),
      "senderLatong": String(senderLatong), //Perbaiki
      "senderSlotStart": senderSlotStart,
      "senderSlotEnd": senderSlotEnd,
      "senderEmail": null,
      "senderAddress": senderAddress,
      "senderNote": null,
      "consigneeName": consigneeName,
      "consigneePhone": consigneePhone,
      "consigneeEmail": null,
      "consigneeDistrict": consigneeDistrict,
      "consigneeCity": consigneeCity,
      "consigneeProvince": consigneeProvince,
      "consigneePostalcode": null,
      "consigneeLonglat": "",
      "consigneeLatlong": "",
      "consigneeSlotStart": Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'HH:mm:ssXXX'),
      "consigneeSlotEnd": Utilities.formatDate(new Date(new Date().getTime() + 3 * 60 * 60 * 1000), 'GMT+7', 'yyyy-MM-dd\'T\'HH:mm:ssXXX'),
      "consigneeAddress": consigneeAddress,
      "consigneeCountry": "Indoneisa",
      "consigneeNote": null,
      "description": null,
      "productId": String(productId),
      "productName": productName,
      "assignee": [email],
      "startTime": Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd\'T\'HH:mm:ssXXX'),
      "endTime": Utilities.formatDate(new Date(new Date().getTime() + 7 * 24 * 60 * 60 * 1000), 'GMT+7', 'yyyy-MM-dd\'T\'HH:mm:ssXXX'),
    };
    tasks.push(task);

    // Mengecek apakah array tasks sudah memiliki elemen atau tidak
    if (tasks.length > 0) {
      // Jika ya, mengirim permintaan POST ke API dengan array tasks sebagai payload
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
    } else {
      // Jika tidak, mengabaikan permintaan POST
      Logger.log("Tidak ada data yang dikirim");
    }
    
    var sheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1zOkoXJKXTfnfc_sn5EaB6PtALOjcD15Ca0vTlUDwB_0/edit#gid=1208648341").getSheetByName("OUPUT_GFORM");
    // Menyimpan nilai faktur pada kolom A lembar kerja "NOTE"
    sheet.appendRow([faktur]); // Menggunakan metode appendRow(faktur) untuk menambahkan nilai faktur pada baris baru
    
    // Menampilkan nilai email pada kolom D lembar kerja "OUPUT_GFORM"
    // Mendapatkan sel pada baris terakhir dan kolom D lembar kerja "OUPUT_GFORM"
    var cell = sheet.getRange(sheet.getLastRow(), 4, 1, 1);
    // Menetapkan nilai email sebagai nilai dari sel tersebut
    cell.setValue(email);
    var cell = sheet.getRange(sheet.getLastRow(), 5, 1, 1);
    // Menetapkan nilai slot sebagai nilai dari sel tersebut
    cell.setValue(slot);
}

function getHubId(cabang) {
  var hubIds = {
    "PT MERAPI UTAMA PHARMA - JK1": "645b1d984c84e849033139f4",
    "PT MERAPI UTAMA PHARMA - JK2": "645b23a4a3e74b10574074e3",
    "PT MERAPI UTAMA PHARMA - DPK": "645b231d211f32fa0802e354",
    "PT MERAPI UTAMA PHARMA - BDG": "645b228073bcfc3f463541e2",
    "PT MERAPI UTAMA PHARMA - BTN": "645b22bfe355136c1234ca47"
  };
  return hubIds[cabang];
}