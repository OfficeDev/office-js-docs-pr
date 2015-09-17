# Build your first Excel add-in

_Applies to: Excel 2016_

An Excel add-in runs inside Excel and can interact with the contents of the spreadsheet using the new Excel JavaScript APIs available in Office 2016. Under the hood, an add-in is simply a web app that you can host anywhere. The add-in manifest (manifest.xml) tells where your web app is located and how you want it to appear within Excel.
  
>**Excel add-in = manifest.xml + your own web app**

You can create two types of Excel add-ins: task pane and content. 

**Task pane add-ins**
Task pane add-ins work side-by-side with the Excel spreadsheet, and let you supply contextual information and functionality to enhance the spreadsheet viewing and authoring experience. For example, a task pane add-in can look up and retrieve product information from a web service based on the product name or part number selected in the document.

**Content add-ins**
Content add-ins integrate web-based features as content that is shown in-line with the contents of a spreadsheet. Content add-ins let you integrate rich, web-based data visualizations, embedded media (such as a YouTube video player or a picture gallery), as well as other external content.

## Your first add-in
The steps below show how to build and run a simple Excel task pane add-in that loads some data into a  worksheet and creates a simple chart. 

### Set it up
1.	To create the add-in, you essentially need to create a web app and an XML manifest file that tells where your web app is located and how you want it to appear within Excel. In this example, you will create the web app using HTML and JQuery. To start, create a folder on your local drive called QuarterlySalesReport (for example C:\QuarterlySalesReport). Save all of the files created in the following steps into this folder.

2.  Let's now create the HTML page that will be loaded into the task pane add-in. Create a file named home.html and paste in the code below.
	```
	
	<!DOCTYPE html>
	<html>
	  <head>
		<meta charset="UTF-8" />
		<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
		<title>Quarterly Sales Report</title>	
		<script src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
		<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>
		<script src="https://officesnippetexplorerbranch.azurewebsites.net/script/Office.runtime.js"></script>
		<script src="https://officesnippetexplorerbranch.azurewebsites.net/script/excel.js"></script>
		<link href="Styles.css" rel="stylesheet" type="text/css" />
		<script src="Home.js" type="text/javascript"></script>
		<link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
		<link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
		</head>
		<body class="ms-font-m">
		    <div id="content-header">
		            <h1>Welcome</h1>
		    </div>
		    <div id="content-main">
		            <p>This sample shows how to load some sample data into the worksheet and then create a chart using the Excel JavaScript API.</p>
		            <br />
		            <h3>Try it out</h3>
		            <button class="ms-Button" id="load-data-and-create-chart">Click me!</button>
		    </div>
		</body>
</html>
	```  
3.    Next create a file named Styles.css to store your custom styles and paste in the code below.
	```
	#content-header {
		    background: #2a8dd4;
		    color: #fff;
		    position: absolute;
		    top: 0;
		    left: 0;
		    width: 100%;
		    height: 80px; /* Fixed header height */
			padding: 10px;
		    overflow: hidden; /* Disable scrollbars for header */
		}
		
	#content-main {
		    background: #fff;
		    position: fixed;
		    top: 80px; /* Same value as #content-header's height */
		    left: 0;
		    right: 0;
		    bottom: 0;
			padding: 15px;
		    overflow: auto; /* Enable scrollbars within main content section */
		}
	
	```

4.  Create a file named Home.js and copy and paste the following script. This file contains the programming logic for the add-in in JQuery.
	```
	(function () {
	    "use strict";
	
	    // The initialize function must be run each time a new page is loaded
	    Office.initialize = function (reason) {
	        $(document).ready(function () {
	            $('#load-data-and-create-chart').click(loadDataAndCreateChart);
	        });
	    };
	
	    // Load some sample data into the worksheet and then create a chart
	    function loadDataAndCreateChart() {
			// Run a batch operation against the Excel object model
	        Excel.run(function (ctx) {
	
	            // Create a proxy object for the active worksheet
	            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
	
	            //Queue commands to set the report title in the worksheet
	            sheet.getRange("A1:A1").values = "Quarterly Sales Report";
	            sheet.getRange("A1:A1").format.font.name = "Century";
	            sheet.getRange("A1:A1").format.font.size = 26;
	
	
	
	            //Create an array containing sample data
	            var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
	                          ["Frames", 5000, 7000, 6544, 4377],
	                          ["Saddles", 400, 323, 276, 651],
	                          ["Brake levers", 12000, 8766, 8456, 9812],
	                          ["Chains", 1550, 1088, 692, 853],
	                          ["Mirrors", 225, 600, 923, 544],
	                          ["Spokes", 6005, 7634, 4589, 8765]];
	
	            //Queue a command to write the sample data to the specified range
				//in the worksheet and bold the header row
	            var range = sheet.getRange("A2:E8");
	            range.values = values;
	            sheet.getRange("A2:E2").format.font.bold = true;
	
	            //Queue a command to add a new chart
	            var chart = sheet.charts.add("ColumnClustered", range, "auto");
	
	            //Queue commands to set the properties and format the chart
	            chart.setPosition("G1", "L10");
	            chart.title.text = "Quarterly sales chart";
	            chart.legend.position = "right"
	            chart.legend.format.fill.setSolidColor("white");
	            chart.dataLabels.format.font.size = 15;
	            chart.dataLabels.format.font.color = "black";
	            var points = chart.series.getItemAt(0).points;
	            points.getItemAt(0).format.fill.setSolidColor("pink");
	            points.getItemAt(1).format.fill.setSolidColor('indigo');
	
	            //Run the queued-up commands, and return a promise to indicate task completion
	            return ctx.sync();
	        })
	          .then(function () {
				console.log("Success!");
	        })
	          .catch(function (error) {
	            console.log("Error: " + error);
	            if (error instanceof OfficeExtension.RuntimeError) {
	                console.log("Debug info: " + JSON.stringify(error.debugInfo));
	            }
	            console.log("Error: " + JSON.stringify(error.debugInfo));
	        }); 
	    }
	})();
	```

5.  Create an XML file named QuarterlySalesReportManifest.xml and copy and paste the following XML. The manifest tells where your web app is located and how you want it to appear within Excel.
	```
	<?xml version="1.0" encoding="UTF-8"?>
	<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
	<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
	  <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
	  <Version>1.0.0.0</Version>
	  <ProviderName>Microsoft</ProviderName>
	  <DefaultLocale>en-US</DefaultLocale>
	  <DisplayName DefaultValue="Quarterly Sales Report" />
	  <Description DefaultValue="Quarterly Sales Report"/>
	  <Capabilities>
	    <Capability Name="Workbook" />
	  </Capabilities>
	  <DefaultSettings>
	    <SourceLocation DefaultValue="\\MyShare\QuarterlySalesReport\Home.html" />
	  </DefaultSettings>
	  <Permissions>ReadWriteDocument</Permissions>
	</OfficeApp>
	```

6.	Generate a GUID using an online generator. Then replace the value in the **Id** tag above with a GUID that you have generated yourself.

7.	Save all the files. You’ve now written your first Excel add-in. 

### Try it out

1.	The simplest way to deploy and test your add-in is to copy the files to a network share. To do this, follow these steps:
	1.  Create a folder on a network share (for example, \\MyShare\QuarterlySalesReport) and copy all the files. 
	2. Edit the <SourceLocation> element of the manifest file so that it points to the share location for the .html page from step 1. 
	3. Then copy the manifest (QuarterlySalesReportManifest.xml) to a network share (for example, \\MyShare\MyManifests).
	4. Then add the share location that contains the manifest as a trusted app catalog in Excel. To do this, follow these steps:
	    1. 	Launch Excel.
	    2. Choose the File tab, and then choose Options.
	    3. Choose Trust Center, and then choose the Trust Center Settings button.
	    4. 	Choose Trusted App Catalogs.
	    5. 	In the Catalog Url box, enter the path to the network share you created in Step 1, and then choose Add Catalog.
	    6. Select the Show in Menu check box, and then choose OK.
	    7. 	A message is displayed to inform you that your settings will be applied the next time you start Office. Close and restart Excel. 
        
2.	Test and run the add-in. To do this, follow these steps:
    1.  On the Insert tab in Excel 2016, choose My Add-ins. 
    2.  In the Office Add-ins dialog box, choose Shared Folder.
    3.  Choose Quarterly Sales Report, and then choose Insert.
    4.  The add-in will open in a task pane to the right of the current worksheet as shown in this diagram. ![Quarterly Sales Report Add-in](/excel/images/QuarterlySalesReport_taskpane.png)
    5. Now click the **Click me!** button. This will render the data and the chart inside the worksheet as shown below.  Feel free to change the data in the range and see the chart get updated dynamically! ![Quarterly Sales Report Add-in](images/QuarterlySalesReport_report.png)

### Learn more

Believe it or not, we’ve only just begun exploring what can be accomplished with the new Excel JavaScript APIs. The APIs have much more to offer. If you’d like to know more, you’re welcome to explore any of the available resources. 

Here are just a few:

1.  [Excel programming guide](excel-add-ins-programming-guide.md)
2.  [Add-in code samples](excel-add-ins-code-samples.md) 
3.  [Reference](excel-add-ins-javascript-reference.md)