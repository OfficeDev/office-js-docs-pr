---
title: Freeze a table header in Excel
description: ''
ms.date: 12/08/2017 
---


# Freeze a table header in place

This is the fourth step of a tutorial that begins with [Excel Tutorial Create Table](excel-tutorial-create-table.md). You need to complete the preceding steps to get the project in the state that this step assumes. 

When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this tutorial, you learn how to freeze a row so that it remains visible even if the user scrolls a great deal. 

## Freeze the table's header row

1. Open the project in your code editor. 
2. Open the file index.html.
3. Below the `div` that contains the `create-chart` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="freeze-header">Freeze Header</button>            
    </div>
    ```

4. Open the app.js file.

5. Below the line that assigns a click handler to the `create-chart` button, add the following code:

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. Below the `createChart` function add the following function:

    ```js
    function freezeHeader() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to keep the header visible when the user scrolls.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. Replace `TODO1` with the following code. Note:
   - The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.
   - The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1`

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ``` 


## Test the add-in

1. Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.
2. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).
3. Run the command `npm start` to start a web server running on localhost.
4. Sideload the add-in using one of the methods described in [Excel add-in quickstart that uses jQuery](../quickstarts/excel-quickstart-jquery.md).
5. On the **Home** menu, choose **Show Taskpane**.
6. In the taskpane, choose **Create Table**. 
7. Choose the **Freeze Header** button.
8. Scroll the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.

    ![Excel tutorial - Freeze Header](../images/excel-tutorial-freeze-header.png)


## Next steps

<div style="display: flex; flex-direction: row; justify-content:space-around">
<a style="padding: 15px 35px; 
	border-radius: 4px; 
	color: rgb(255, 255, 255); 
	line-height: 3em; 
	font-family: wf_segoe-ui_semibold,wf_segoe-ui_normal,Helvetica Neue,Helvetica,Arial,sans-serif; 
	font-size: 16px; 
	font-weight: 400; 
	margin-left: 20px; 
	white-space: nowrap; 
	cursor: pointer; 
	background-color: rgb(0, 120, 215);
}" href="excel-tutorial-protect-worksheet.md">I'm ready for the <b>Protect Worksheet</b> tutorial</a>     
<a style="padding: 13px 33px; 
	border-radius: 4px; 
	border: 2px solid rgb(0, 120, 215); 
	border-image: none; 
	color: rgb(0, 120, 215); 
	line-height: 3em; 
	font-family: wf_segoe-ui_semibold,wf_segoe-ui_normal,Helvetica Neue,Helvetica,Arial,sans-serif; 
	font-size: 16px; 
	font-weight: 400; 
	white-space: nowrap; 
	cursor: pointer;" href="https://github.com/OfficeDev/office-js-docs/issues">I ran into an issue</a>
          
</div>

