function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generuj Reporty')
    .addItem('Generuj Analyzy', 'GenerateAll')
    .addToUi();
}

function GenerateAll() {
  const analyses = getAllAnalyses()
  GenerateAnalysesAll(analyses)
  GenerateAnalysesPerSwimmer(analyses)
}

// Generate analyses for all swimmers
function GenerateAnalysesAll(analyses) {
  let analysesJSON = {
    "name": "all",
    "reports": []
  }

  const names = analyses.forEach(
    row => {
      if (row[0] != "") {
        analysesJSON.reports.push(getReportDict(row))
      }
    }
  )

  const baseFolder = getFolder(getParentFolder(), 'rozbory')
  downloadReports(baseFolder, 'https://plaway-docs.ey.r.appspot.com/analyses', analysesJSON)
}


// Generate per swimmer analyses
function GenerateAnalysesPerSwimmer(analyses) {
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
        analysesJSON[row[1]].reports.push(getReportDict(row))
      }
    }
  )

  const baseFolder = getFolder(getParentFolder(), 'rozbory')
  for (const key in analysesJSON) {
    const swimmerReport = analysesJSON[key];
    const swimmerFolder = getFolder(baseFolder, swimmerReport.name)
    downloadReports(swimmerFolder, 'https://plaway-docs.ey.r.appspot.com/analyses', swimmerReport)
  }
}


function getReportDict(row) {
  return {
    "recommendedDrills": row[8],
    "toImprove": row[7],
    "whatsGreat": row[6],
    "souhra": row[5],
    "legs": row[4],
    "arms": row[3],
    "style": row[2],
    "name": row[1]}
}

// Return all analyses in report
function getAllAnalyses() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA')
    .getDataRange()
    .getValues()
    .slice(2);
}

function downloadReports(destinationFolder: GoogleAppsScript.Drive.Folder, baseURL: string, data) {
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
    createUniqueFileInFolder(destinationFolder, fileBlob, data.name + ".pdf")
  }
}

// Return empty folder relative to this sheet
function getFolder(baseFolder: GoogleAppsScript.Drive.Folder, name: string) {
  const folders = baseFolder.getFolders()
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName() == name) {
      Logger.log("Returning existing folder " + name + " in " + baseFolder.getName())
      return folder
    }
  }

  Logger.log("Creating folder " + name + " in " + baseFolder.getName())
  return baseFolder.createFolder(name)
}

function createUniqueFileInFolder(folder: GoogleAppsScript.Drive.Folder, fileBlob: GoogleAppsScript.Base.Blob, filename: string) {
  const folderName = folder.getName()

  Logger.log("Creating file " + filename + " in folder " + folderName)
  const files = folder.getFiles()
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName() == filename) {
      Logger.log("Deleting existing file " + filename + " in folder" + folderName)
      file.setTrashed(true)
    }
  }

  fileBlob.setName(filename)
  folder.createFile(fileBlob)
}

function getParentFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileInDrive = DriveApp.getFolderById(ss.getId());

  return fileInDrive.getParents().next()
}
