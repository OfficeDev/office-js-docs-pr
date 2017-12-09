---
title: Create a chart in Excel
description: ''
ms.date: 12/08/2017 
---


# Create a chart

This is the third step of a tutorial that begins with [Excel Tutorial Create Table](excel-tutorial-create-table.md). You need to complete the preceding steps to get the project in the state that this step assumes. 

In this tutorial, you'll learn how to programmatically create a chart from table data and how to format the chart. 

## Chart table data

1. Open the project in your code editor. 
2. Open the file index.html.
3. Below the `div` that contains the `sort-table` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-chart">Create Chart</button>            
    </div>
    ```

4. Open the app.js file.

5. Below the line that assigns a click handler to the `sort-chart` button, add the following code:

    ```js
    $('#create-chart').click(createChart);
    ```

6. Below the `sortTable` function add the following function.

    ```js
    function createChart() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

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

7. Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ``` 

8. Replace `TODO2` with the following code. Note the following parameters:
   - The first parameter to the `add` method specifies the type of chart. There are several dozen types. 
   - The second parameter specifies the range of data to include in the chart. 
   - The third parameter determines whether a series of data points from the table should be charted rowwise or columnwise. The option `auto` tells Excel to decide the best method.

    ```js
    let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ``` 

9. Replace `TODO3` with the following code. Most of this code is self-explanatory. Note:
   - The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.
   - A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in €". 

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ``` 

## Test the add-in

1. Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.
3. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used by Excel to run Excel add-ins).
4. Run the command `npm start` to start a web server running on localhost.
5. Sideload the add-in using one of the methods described in [Excel add-in quickstart that uses jQuery](excel-add-ins-get-started-jquery.md).
6. On the **Home** menu, select **Show Taskpane**.
7. In the taskpane, choose **Create Table**. 
8. Choose the **Filter Table** and **Sort Table** buttons, in either order.
9. Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.


    ![Excel tutorial - Create Chart](../images/excel-tutorial-create-chart.png)



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
}" href="excel-tutorial-freeze-header.md">I'm ready for the <b>Freeze Header</b> tutorial</a>     
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

