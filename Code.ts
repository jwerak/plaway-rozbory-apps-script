function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Generuj Reporty')
      .addItem('Export JSON', 'exportJSONData')
      .addToUi();
}


function exportJSONData(){
  let analyses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA')
    .getDataRange()
    .getValues()
    .slice(2);

  let analysesJSON = {
    "name": "all",
    "reports": []
  }

  const names = analyses.forEach(

    row => {
      if (row[0] != "") {
        analysesJSON.reports.push(
        {
          "recommendedDrills": row[8],
          "toImprove": row[7],
          "whatsGreat": row[6],
          "souhra": row[5],
          "legs": row[4],
          "arms": row[3],
          "style": row[2],
          "name": row[1]}
        )
      }
    }
  )

  downloadReports('https://plaway-docs.ey.r.appspot.com/analyses', analysesJSON)
}


function downloadReports(baseURL, data) {
  let fileName = "";
  let fileSize = 0;
  var options = {
    muteHttpExceptions: true,
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization' : 'Bearer epheulieNgiexiLodi4lai2Bim5ieshowahdooCa'},
    payload: JSON.stringify(data),
  };
  const response = UrlFetchApp.fetch(baseURL, options);
  const rc = response.getResponseCode();

  if (rc == 200) {
    const fileBlob = response.getBlob()
    fileBlob.setName("analyzy.pdf")
    DriveApp.createFile(fileBlob)
    // const file = folder.createFile(fileBlob);
    // fileName = file.getName();
    // fileSize = file.getSize();
  }

  // var fileInfo = { "rc":rc, "fileName":fileName, "fileSize":fileSize };
  // return fileInfo;
}
