function ExportWorklogAsExcel() {
    try {
        var ss = SpreadsheetApp.getActive();
        var availableSheets = ss.getSheets().map(sheet => sheet.getName())
        var settingSheet = ss.getSheetByName("Experiment")
        var sendToEmails = settingSheet.getRange("B2").getDisplayValue().split(",").map(em => em.trim())
        var exportedMonths = settingSheet.getRange("B3").getDisplayValue().split(",").map(mo => mo.trim())
        var currentSheet = ss.getActiveSheet()

        exportedMonths.forEach(month => {
            let exportedSheets = availableSheets.filter(sheetName => sheetName.endsWith(" - " + month))
            let unexportedSheets = availableSheets.filter(sheetName => !sheetName.endsWith(" - " + month))
            exportedSheets.forEach(sheetName => {
                ss.getSheetByName(sheetName).showSheet()
            })
            unexportedSheets.forEach(sheetName => {
                ss.getSheetByName(sheetName).hideSheet()
            })

            if (exportedSheets.length > 0) {
              fileName = `PTTEP Worksheet - ${month}.xlsx`
                sendTo = sendToEmails.join(",")
                mailSubject = `PTTEP Worksheet - ${month}`
                mailBody = "Please see the attachment."
                
                SendCurrentSheetAsExcelTo_(fileName, sendTo, mailSubject, mailBody)
                Logger.log(`Worklog for ${month} has been exported and sent to ${sendTo}`)
            }

            // Show the hidden
            unexportedSheets.forEach(sheetName => {
                ss.getSheetByName(sheetName).showSheet()
            })
        })

        currentSheet.activate()
    } catch (f) {
        Logger.log(f.toString());
    }
}

function SendCurrentSheetAsExcelTo_(excelFileName, sendTo, mailSubject, mailBody) {
    try {
        var ss = SpreadsheetApp.getActive();
        var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + ss.getId() + "&exportFormat=xlsx";
        var params = {
            method: "get",
            headers: {
                "Authorization": "Bearer " + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
        };
        var blob = UrlFetchApp.fetch(url, params).getBlob();
        blob.setName(excelFileName);
        MailApp.sendEmail(sendTo, mailSubject, mailBody, {
            attachments: [blob]
        });

    } catch (f) {
        Logger.log(f.toString());
    }
}