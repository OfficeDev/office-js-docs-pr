---
title: Filter and sort a table in Excel
description: ''
ms.date: 12/08/2017 
---


# Filter and Sort a Table

This is the second step of a tutorial that begins with [Excel Tutorial Create Table](excel-tutorial-create-table.md). You need to complete the preceding steps to get the project in the state that this step assumes. 

This tutorial shows you how to programmatically filter and sort a table.

## Filter the table

1. Open the project in your code editor. 
2. Open the file index.html.
3. Just below the `div` that contains the `create-table` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="filter-table">Filter Table</button>            
    </div>
    ```

4. Open the app.js file.

5. Just below the line that assigns a click handler to the `create-table` button, add the following code:

    ```js
    $('#filter-table').click(filterTable);
    ```

6. Just below the `createTable` function add the following function:

    ```js
    function filterTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to filter out all expense categories except 
            //        Groceries and Education.

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
   - The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.
   - The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## Sort the table

1. Open the file index.html.
2. Below the `div` that contains the `filter-table` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="sort-table">Sort Table</button>            
    </div>
    ```

3. Open the app.js file.

4. Below the line that assigns a click handler to the `filter-table` button, add the following code:

    ```js
    $('#sort-table').click(sortTable);
    ```

5. Below the `filterTable` function add the following function.

    ```js
    function sortTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to sort the table by Merchant name.

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
   - The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.
   - The `key` property of a `SortField` object is the zero-based index of the column to sort-on.
   - The `sort` member of a `Table` is a `TableSort` object, not a method. The `SortField`s are passed the `TableSort` object's `apply` method.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
        { 
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ``` 

## Test the add-in

1. Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.
2. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).
3. Run the command `npm start` to start a web server running on localhost.
4. Sideload the add-in using one of the methods described in [Excel add-in quickstart that uses jQuery](../quickstarts/excel-quickstart-jquery.md).
5. On the **Home** menu, select **Show Taskpane**.
6. In the taskpane, choose **Create Table**. 
7. Choose the **Filter Table** and **Sort Table** buttons, in either order.

    ![Excel tutorial - Filter and Sort Table](../images/excel-tutorial-filter-and-sort-table.png)


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
}" href="excel-tutorial-create-chart.md">I'm ready for the <b>Create Chart</b> tutorial</a>     
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




