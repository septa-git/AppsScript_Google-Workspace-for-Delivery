function onFormSubmit(event) {
  var record_array = [];

  var form = FormApp.openById('12CIe4_SRrXh62ac31kzY4Q9KdxcGA2cBjdkpRR1GRa8'); // Form ID
  var formResponses = form.getResponses();
  var formCount = formResponses.length;

  var formResponse = formResponses[formCount - 1];
  var itemResponses = formResponse.getItemResponses();

  for (var j = 0; j < itemResponses.length; j++) {
    var itemResponse = itemResponses[j];
    var title = itemResponse.getItem().getTitle();
    var answer = itemResponse.getResponse();

    Logger.log(title);
    Logger.log(answer);

    // Jika jawaban adalah array (misalnya dari checkbox), gabungkan menjadi string
    if (Array.isArray(answer)) {
      answer = answer.join(", ");
    }

    record_array.push(answer);
  }
   
  AddRecord(record_array);
}

function AddRecord(record_array) {
  var url = 'https://docs.google.com/spreadsheets/d/1zOkoXJKXTfnfc_sn5EaB6PtALOjcD15Ca0vTlUDwB_0/edit#gid=1740099260';   //URL OF GOOGLE SHEET;
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("GET_INVOICES");
  
  // Menambahkan tanggal dan waktu saat ini ke array
  record_array.push(new Date());
  
  // Menambahkan baris baru ke lembar data dengan semua nilai input pengguna
  dataSheet.appendRow(record_array);
}