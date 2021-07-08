---
title: Work with workbooks using the Excel JavaScript API
description: 'Learn how to perform common tasks with workbooks or application-level features using the Excel JavaScript API.'
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
---

# Work with workbooks using the Excel JavaScript API

This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API. For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook). This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.

The Workbook object is the entry point for your add-in to interact with Excel. It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed. The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets. Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events. The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.

## Get the active cell or selected range

The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`. `getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range). The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

The `getSelectedRange()` method returns the currently selected single range. If multiple ranges are selected, an InvalidSelection error is thrown. The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## Create a workbook

Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running. The Excel object has the `createWorkbook` method for this purpose. When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel. Your add-in remains open and running with the previous workbook.

```js
Excel.createWorkbook();
```

The `createWorkbook` method can also create a copy of an existing workbook. The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter. The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.

You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### Insert a copy of an existing workbook into the current one

The previous example shows a new workbook being created from an existing workbook. You can also copy some or all of an existing workbook into the one currently associated with your add-in. A [Workbook](/javascript/api/excel/excel.workbook) has the `insertWorksheetsFromBase64` method to insert copies of the target workbook's worksheets into itself. The other workbook's file is passed as a base64-encoded string, just like the `Excel.createWorkbook` call. 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> The `insertWorksheetsFromBase64` method is supported for Excel on Windows, Mac, and the web. It's not supported for iOS. Additionally, in Excel on the web this method doesn't support source worksheets with PivotTable, Chart, Comment, or Slicer elements. If those objects are present, the `insertWorksheetsFromBase64` method returns the `UnsupportedFeature` error in Excel on the web. 

The following code sample shows how to insert worksheets from another workbook into the current workbook. This code sample first processes a workbook file with a [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) object and extracts a base64-encoded string, and then it inserts this base64-encoded string into the current workbook. The new worksheets are inserted after the worksheet named **Sheet1**. Note that `[]` is passed as the parameter for the [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) property. This means that all the worksheets from the target workbook are inserted into the current workbook.

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## Protect the workbook's structure

Your add-in can control a user's ability to edit the workbook's structure. The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method. The following example shows a basic scenario toggling the protection of the workbook's structure.

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

The `protect` method accepts an optional string parameter. This string represents the password needed for a user to bypass protection and change the workbook's structure.

Protection can also be set at the worksheet level to prevent unwanted data editing. For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.

> [!NOTE]
> For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.

## Access document properties

Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values. The following example shows how to set the `author` property.

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### Custom properties

You can also define custom properties. The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties. The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### Worksheet-level custom properties

Custom properties can also be set at the worksheet level. These are similar to document-level custom properties, except that the same key can be repeated across different worksheets. The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## Access document settings

A workbook's settings are similar to the collection of custom properties. The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file. The following example shows how to create and access a setting.

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## Access application culture settings

A workbook has language and culture settings that affect how certain data is displayed. These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures. Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.

`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object. This contains settings like the numerical decimal separator or the date format.

Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e). The system settings are preserved in the `CultureInfo` object. Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.

The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## Add custom XML data to the workbook

Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook. This data persists with the workbook, independent of the add-in.

A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart). These give access to the XML strings and a corresponding unique ID. By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.

The following samples show how to use custom XML parts. The first code block demonstrates how to embed XML data in the document. It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval. The second block shows how to access that XML later. The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`. The XML data is then printed to the console.

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> `CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.

## Control calculation behavior

### Set calculation mode

By default, Excel recalculates formula results whenever a referenced cell is changed. Your add-in's performance may benefit from adjusting this calculation behavior. The Application object has a `calculationMode` property of type `CalculationMode`. It can be set to the following values.

- `automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.
- `automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.
- `manual`: Calculations only occur when the user or add-in requests them.

### Set calculation type

The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation. `Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`. The following values can be specified.

- `full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.
- `fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.
- `recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.

> [!NOTE]
> For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.

### Temporarily suspend calculations

The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called. This is done with `suspendApiCalculationUntilNextSync()`. Use this method when your add-in is editing large ranges without needing to access the data between edits.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## Detect workbook activation

Your add-in can detect when a workbook is activated. A workbook becomes *inactive* when the user switches focus to another workbook, to another application, or (in Excel on the web) to another tab of the web browser. A workbook is *activated* when the user returns focus to the workbook. The workbook activation can trigger callback functions in your add-in, such as refreshing workbook data.

To detect when a workbook is activated, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onActivated](/javascript/api/excel/excel.workbook#onActivated) event of a workbook. Event handlers for the `onActivated` event receive a [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) object when the event fires.

> [!IMPORTANT]
> The `onActivated` event doesn't detect when a workbook is opened. This event only detects when a user switches focus back to an already open workbook.

The following code sample shows how to register the `onActivated` event handler and set up a callback function.

```js
Excel.run(function (context) {
    // Retrieve the workbook.
    var workbook = context.workbook;

    // Register the workbook activated event handler.
    workbook.onActivated.add(workbookActivated);

    return context.sync();
});

function workbookActivated(event) {
    Excel.run(function (context) {
        // Retrieve the workbook and load the name.
        var workbook = context.workbook;
        workbook.load("name");
        
        return context.sync().then(function () {
            // Callback function for when the workbook is activated.
            console.log(`The workbook ${workbook.name} was activated.`);
        });
    });
}
```

## Save the workbook

`Workbook.save` saves the workbook to persistent storage. The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values.

- `Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location. If the file has not been saved previously, it's saved to the default location. If the file has been saved previously, it's saved to the same location.
- `Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location. If the file has been saved previously, it will be saved to the same location and the user will not be prompted.

> [!CAUTION]
> If the user is prompted to save and cancels the operation, `save` throws an exception.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## Close the workbook

`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open). The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values.

- `Excel.CloseBehavior.save` (default): The file is saved before closing. If the file has not been saved previously, the user will be prompted to specify file name and save location.
- `Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving. Any unsaved changes will be lost.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
