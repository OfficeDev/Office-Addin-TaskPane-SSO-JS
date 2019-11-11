export function writeFileNamesToOfficeDocument(result) {

    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            switch (Office.context.host) {
                case Office.HostType.Excel:
                    writeFileNamesToWorksheet(result);
                    break;
                case Office.HostType.Word:
                    writeFileNamesToDocument(result);
                    break;
                case Office.HostType.PowerPoint:
                    writeFileNamesToPresentation(result);
                    break;
                default:
                    throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
            }
            resolve();
        }
        catch (error) {
            reject(Error("Unable to add filenames to document. " + error.toString()));
        }
    });
}

function writeFileNamesToWorksheet(result) {

    return Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        let filenames = [];
        let i;
        for (i = 0; i < result.length; i++) {
            var innerArray = [];
            innerArray.push(result[i]);
            filenames.push(innerArray);
        }

        const rangeAddress = `B5:B${5 + (result.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = filenames;
        range.format.autofitColumns();

        return context.sync();
    });
}

function writeFileNamesToDocument(result) {
    return Word.run(function (context) {
        const documentBody = context.document.body;
        for (let i = 0; i < result.length; i++) {
            documentBody.insertParagraph(result[i], "End");
        }

        return context.sync();
    });
}

function writeFileNamesToPresentation(result) {

    let fileNames = "";
    for (let i = 0; i < result.length; i++) {
        fileNames += result[i] + '\n';
    }

    Office.context.document.setSelectedDataAsync(
        fileNames,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
            }
        }
    );
}