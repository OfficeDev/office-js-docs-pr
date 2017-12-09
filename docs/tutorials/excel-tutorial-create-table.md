---
title: Create a table in Excel
description: ''
ms.date: 12/08/2017 
---


# Create and populate a table

This is the first step in a series of tutorials. Each one adds to the same project. 

>**Note**: If you haven't already, please read [Excel add-in quickstart that uses jQuery](../quickstarts/excel-quickstart-jquery.md). In particular, be sure that you know how to sideload an Excel add-in for testing.

This first tutorial shows you how to programmatically add a table to a worksheet, populate the table with data, and then format it. It also shows you how to test that your add-in supports the user's current version of Excel.


## Prerequisites

To use this tutorial, you need to have the following installed. 

- Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).
- [Node and npm](https://nodejs.org/en/) 
- [Git Bash](https://git-scm.com/downloads) (Or another git client.)

## Setup

1. Clone the repo [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).
2. Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.
3. Run the command `npm install` to install the tools and libraries listed in the package.json file. 

## Code the add-in

1. Open the project in your code editor. 
2. Open the file index.html.
3. Replace the `TODO1` with the following markup:

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. Open the app.js file.
5. Replace the `TODO1` with the following code. This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    } 
    ```

6. Replace the `TODO2` with the following code:

    ```js
    $('#create-table').click(createTable);
    ```

7. Replace the `TODO3` with the following code. Note the following:
   - Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.
   - The `context.sync` method sends all queued commands to Excel for execution.
   - The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow. 

    ```js
    function createTable() {
        Excel.run(function (context) {
            
            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. Replace `TODO4` with the following code. Note:
   - The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object. 
   - The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created. 
   - Table names must be unique across the entire workbook, not just the worksheet.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ``` 

9. Replace `TODO5` with the following code. Note:
   - The cell values of a range are set with an array of arrays.
   - New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.

    ```js
    expensesTable.getHeaderRowRange().values = 
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ``` 

10. Replace `TODO6` with the following code. Note:
   - The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection. 

     >**Note**: Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.

   - The code then formats the range of the **Amount** column as Euros to the second decimal. 
   - Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ``` 

## Test the add-in

1. Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.
3. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).
4. Run the command `npm start` to start a web server running on localhost.   
5. Sideload the add-in using one of the methods specified in [Excel add-in quickstart that uses jQuery](../quickstarts/excel-quickstart-jquery.md).
6. On the **Home** menu, choose **Show Taskpane**.
7. In the taskpane, choose **Create Table**. 


    ![Excel tutorial - Create Table](../images/excel-tutorial-create-table.png)


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
}" href="excel-tutorial-filter-and-sort-table.md">I'm ready for the <b>Filter and Sort a Table</b> tutorial</a>     
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

