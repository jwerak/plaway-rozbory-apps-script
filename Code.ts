function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generuj Reporty')
    .addItem('Generuj Analyzy dohromady', 'GenerateAnalysesAll')
    .addItem('Generuj Analyzy zvlast pro plavce', 'GenerateAnalysesPerSwimmer')
    .addToUi();
}


// Generate analyses for all swimmers
function GenerateAnalysesAll() {
  const analyses = getAllAnalyses()

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


// Generate per swimmer analyses
function GenerateAnalysesPerSwimmer(){
  const analyses = getAllAnalyses()

  let analysesJSON = {}

  const names = analyses.forEach(
    row => {
      if (row[0] != "") {
        // Is swimmer missing?
        if (!(row[1] in analysesJSON)) {
          analysesJSON[row[1]] = {
            "name": row[1],
            "reports": []
          }
        }

        // Add analysis
        analysesJSON[row[1]].reports.push(
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
  for (const key in analysesJSON) {
    const swimmerReport = analysesJSON[key];
    downloadReports('https://plaway-docs.ey.r.appspot.com/analyses', swimmerReport)
  }
}

// Return all analyses in report
function getAllAnalyses() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA')
    .getDataRange()
    .getValues()
    .slice(2);
}

function downloadReports(baseURL: string, data) {
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
    fileBlob.setName(data.name + ".pdf")
    DriveApp.createFile(fileBlob)
  }
}

// Return empty folder relative to this sheet
function getEmptyFolderForSwimmer(name: string) {

}
