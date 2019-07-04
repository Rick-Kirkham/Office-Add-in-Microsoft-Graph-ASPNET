// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

Office.initialize = function () {
    $(document).ready(function () {
        app.initialize();

        $("#getOneDriveFilesButton").click(getFileNamesFromGraph);
        $("#logoutO365PopupButton").click(function () {
            window.location.href = "/azureadauth/logout";
        });        
    });
};

function getFileNamesFromGraph() {

    $("#instructionsContainer").hide();
    $("#waitContainer").show();

    $.ajax({
        url: "/files/onedrivefiles",
        type: "GET"
    })
        .done(function (result) {
            writeFileNamesToOfficeDocument(result)
                .then(function (value) {
                    $("#waitContainer").hide();
                    $("#finishedContainer").show();
                })
                .catch(function (error) {
                    console.log(error);
                });
        })
        .fail(function (result) {
            throw "Cannot get data from MS Graph: " + result;
        });
}

function writeFileNamesToOfficeDocument(result) {

    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            switch (Office.context.host) {
                case "Excel":
                    writeFileNamesToWorksheet(result);
                    break;
                case "Word":
                    writeFileNamesToDocument(result);
                    break;
                case "PowerPoint":
                    writeFileNamesToPresentation(result);
                    break;
                default:
                    throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
            }
            resolve();
        }
        catch (error) {
            reject(Error("Unable to add filenames to document. " + error));
        }
    });    
}

function writeFileNamesToWorksheet(result) {
    
     return Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const data = [
             [result[0]],
             [result[1]],
             [result[2]]];

        const range = sheet.getRange("B5:B7");
        range.values = data;
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

    const fileNames = result[0] + '\n' + result[1] + '\n' + result[2];

    Office.context.document.setSelectedDataAsync(
        fileNames,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
            }
        }
    );
}