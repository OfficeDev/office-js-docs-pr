---
ms.date: 04/30/2019
description: Create custom functions in Excel using JavaScript.
title: Create custom functions in Excel (preview)
localization_priority: Priority
---

# Create custom functions in Excel (preview)

Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. This article describes how to create custom functions in Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet. The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

The following code defines the `ADD42` custom function.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.

## Components of a custom functions add-in project

If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall. We'll concentrate on the files that are important to custom functions: 

| File | File format | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>or<br/>**./src/functions/functions.ts** | JavaScript<br/>or<br/>TypeScript | Contains the code that defines custom functions. |
| **./src/functions/functions.html** | HTML | Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions. |
| **./manifest.xml** | XML | Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table. It also lists the locations of other files your add-in might make use of, such as the task pane files and command files. |

### Script file

The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.

The following code defines the custom function `add`  and then specifies association information for the function. For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).

The following code also provides code comments which define the function. The required `@customfunction` comment is declared first, to indicate that this is a custom function. Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties. Finally, a `returns` description is given. For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions. Projects prepared with the current version of the Yo Office generator reference the correct CDN. If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### Manifest file

The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files. 

The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions. If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).

> [!NOTE] 
> The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in. A namespace can only contain alphanumeric characters and periods.

## Saving and sharing state

Custom functions can save data in global JavaScript variables, which can be used in subsequent calls. Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.

The following code sample shows an implementation of a temperature-streaming function that saves state globally. Note the following about this code:

- The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.

- Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.

- If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs. 

- The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable. Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## Coauthoring

Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions. If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in. Once you both have loaded the add-in, the custom function will share results through coauthoring.

For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## Working with ranges of data

Your custom function may accept a range of data as an input parameter, or it may return a range of data. In JavaScript, a range of data is represented as a two-dimensional array.

For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## Determine which cell invoked your custom function

In some cases you'll need to get the address of the cell that invoked your custom function. This may be useful in the following types of scenarios:

- Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.
- Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.
- Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.

The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file. The following sample gives an example of this if you were to write this JSON file by hand. You can also use the `@requiresAddress` tag if automatically generating your JSON file. For more details, see [JSON Autogeneration](custom-functions-json-autogeneration.md).

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

In the script file (**./src/functions/functions.js** or **./src/functions/functions.ts**), you'll also need to add a `getAddress` function to find a cell's address. This function may take parameters, as shown in the following sample as `parameter1`. The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`. For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.

## Known issues

See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions debugging](custom-functions-debugging.md)
