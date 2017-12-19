---
title: Convert an Office Add-in task pane template in Visual Studio to TypeScript
description: ''
ms.date: 12/04/2017
---

# Convert an Office Add-in task pane template in Visual Studio to TypeScript


You can use the Office Add-in JavaScript template in Visual Studio to create an add-in that uses TypeScript. After you create the new add-in in Visual Studio, you can convert the project to TypeScript.  That way, you don't have to start the Office Add-in TypeScript project from scratch.  

> [!NOTE]
> To learn how to create an Office Add-in TypeScript project without using Visual Studio, see  [Create an Office Add-in using any editor](../tutorials/create-an-office-add-in-using-any-editor.md).

In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript. 

This article shows you how to convert an Excel add-in task pane template in Visual Studio from JavaScript to TypeScript. You can use the same steps to convert other Office Add-in JavaScript templates to TypeScript.

To view or download the code sample that this article is based on, see [Excel-Add-In-TS-Start](https://github.com/OfficeDev/Excel-Add-In-TS-Start) on GitHub.

## Prerequisites

Make sure that you have the following installed:

* [Visual Studio 2015 or later](https://www.visualstudio.com/downloads/)
* [Office Developer Tools for Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* [Cumulative Servicing Release for Microsoft Visual Studio 2015 Update 3 (KB3165756)](https://msdn.microsoft.com/en-us/library/mt752379.aspx)
* Excel 2016
* [TypeScript 2.1 for Visual Studio 2015](http://download.microsoft.com/download/6/D/8/6D8381B0-03C1-4BD2-AE65-30FF0A4C62DA/TS2.1-dev14update3-20161206.2/TypeScript_Dev14Full.exe) (after you install Visual Studio 2015 Update 3)

> [!NOTE]
> For more information about installing TypeScript 2.1, see [Announcing TypeScript 2.1](https://blogs.msdn.microsoft.com/typescript/2016/12/07/announcing-typescript-2-1/).

## Create new add-in project

1.  Open Visual Studio and go to **File** > **New** > **Project**. 
2.  Under **Office/SharePoint**, choose **Excel Add-in** and then choose **OK**.

	![Visual Studio Excel Add-in template](../images/visual-studio-addin-template.png)

3.  In the app creation wizard, choose **Add new functionalities to Excel** and choose **Finish**.
4.  Do a quick test of the newly created Excel add-in by pressing F5 or choosing the **Start** button to launch the add-in. The add-in will be hosted locally on IIS, and Excel will open with the add-in loaded.

## Convert the add-in project to TypeScript

1. In **Solution Explorer**, change the Home.js file to Home.ts.
2. Select **Yes** when asked if you're sure you want to change file name extension.  
3. Select **Yes** when asked if you want to search for TypeScript typings search on nuget, as shown in the following screenshot. This opens the **Nuget Package Manager**.

	![Search for TypeScript typings dialog box](../images/search-typescript-typings.png)

4. Choose **Browse** in the **Nuget Package Manager**.  
5. In the search box, type **office-js tag:typescript**.
6. Install **office.js.TypeScript.DefinitelyTyped** and **jquery.TypeScript.DefinitelyTyped**, as shown in the following screenshot.

	![TypeScript DefinitelyTyped NuGets](../images/typescript-definitely-typed-nugets.png)

7. Open Home.ts (formerly Home.js). Remove the following reference from the top of the Home.ts file:

	```javascript
	///<reference path="/Scripts/FabricUI/MessageBanner.js" />
	```

8. Add the following declaration at the top of the Home.ts file:

	```javascript
	declare var fabric: any;
	```

9. Change **‘1.1’** to **1.1**; that is, remove the quotes from the following line in the Home.ts file:

	```javascript
	if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
	```
 
## Run the converted add-in project

1. Press F5 or choose the **Start** button to launch the add-in. 
2. After Excel launches, press the **Show Taskpane** button on the **Home** ribbon.
3. Select all the cells with numbers.
4. Press the **Highlight** button on the task pane. 

## Home.ts code file

For your reference, the following is the code included in the Home.ts file. This file includes the minimum number of changes needed in order for your add-in to run.

> [!NOTE]
> For a complete example of a JavaScript file that has been converted to TypeScript, see [Excel-Add-In-TS-StartWeb/Home.ts](https://github.com/OfficeDev/Excel-Add-In-TS-Start/blob/master/Excel-Add-In-TS-StartWeb/Home.ts). 

```javascript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
                $("#template-description").text("This sample will display the value of the cells you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(
                    displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(
                hightlightHighestValue);
        });
    }

    function loadSampleData() {

        var values = [
                        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
                        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
                        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {

        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {

            // Create a proxy object for the selected range and load its address and values properties.
            var sourceRange = ctx.workbook.getSelectedRange().load("values, address, rowIndex, columnIndex, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync().
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    cellToHighlight.load("values");
                })
                   // Run the queued-up commands.
                .then(ctx.sync)
                .then(function () {
                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync)
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors.
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
```


## See also

* [Promise implementation discussion on StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [Office Add-in samples on GitHub](https://github.com/officedev)
