---
ms.date: 09/20/2018
description: Create a custom function in Excel using JavaScript. 
title: Create custom functions in Excel (Preview)
---

# Create custom functions in Excel (Preview)

Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any other native function in Excel, such as `SUM()`. This article describes how to create custom functions in Excel.

The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet. The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

The following code defines the `ADD42` custom function.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online. To try them, complete these steps:

1. Install Office (build 10827 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program. You must join the Office Insider program in order to have access to custom functions; currently, custom functions are disabled across all Office builds unless you are a member of the Office Insider program.

2. Use [Yo Office](https://github.com/OfficeDev/generator-office) to create an Excel Custom Functions add-in project, and then follow the instructions in the [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to use the project.

3. Type `=CONTOSO.ADD42(1,2)` into any cell of an Excel worksheet, and press **Enter** to run the custom function.

> [!NOTE]
> The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.

## Learn the basics

In the custom functions project that you've created using [Yo Office](https://github.com/OfficeDev/generator-office), you’ll see the following files:

| File | File format | Description |
|------|-------------|-------------|
| **./src/customfunctions.js** | JavaScript | Contains the code that defines custom functions. |
| **./config/customfunctions.json** | JSON | Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users. |
| **./index.html** | HTML | Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions. |
| **./manifest.xml** | XML | Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table. |

### Manifest file (./manifest.xml)

The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files. The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function `ADD42()` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in. 

### JSON file (./config/customfunctions.json)

A custom functions metadata file provides the information that Excel requires to register the custom functions and make them available to end-users. Custom functions are registered when a user runs an add-in for the first time. After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)

> [!TIP]
> Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.

The following code in **customfunctions.json** specifies the metadata for the `ADD42` function that was described previously in this article. This metadata defines the function's name, description, return value, input parameters, and more. The table that follows this code sample provides detailed information about the individual properties within this JSON object.

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://www.contoso.com/help",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
        }
    ]
}
```

The following table lists the properties that are typically present in the JSON metadata file. For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).

| Property 	| Description |
|---------|---------|
| `id` | A unique ID for the function. This ID should not be changed after it is set. |
| `name` | Name of the function that is shown in the autocomplete menu as a user types a formula within a cell. In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file. |
| `helpUrl`	| Url for a page that is shown when a user requests help. |
| `description`	| Describes what the function does. This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel. |
| `result` 	| Object that defines the type of information that is returned by the function. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`). |
| `parameters` | Array that defines the input parameters for the function. The `name` and `description` child properties appear in the Excel intelliSense. The `type` and `dimensionality` child properties are identical to the child properties of the `result` object that is described previously in this table. |
| `options`	| Enables you to customize some aspects of how and when Excel executes the function. For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article. |

## Functions that return data from external sources

If a custom function retrieves data from an external source such as the web, it must:

1. Return a JavaScript Promise to Excel.

2. Resolve the Promise with the final value using the callback function.

Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result. Users can interact normally with the rest of the worksheet while they wait for the result.

In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer. Note that `sendWebRequest` is a hypothetical function (not specified here) that uses XHR to call a temperature web service.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## Streamed functions

Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation. The following code sample is a custom function that adds a number to the result every second. Note the following about this code:

- Excel displays each new value automatically using the `setResult` callback.

- The final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function. It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.

- In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the JSON metadata file.

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## Canceling a function

In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load. Excel cancels the execution of a function in the following situations:

- When the user edits or deletes a cell that references the function.

- When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.

- The user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.

> [!NOTE]
> You must implement a cancellation handler for every streaming function.

To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the JSON metadata file.

The following code shows the same `incrementValue` function that was described previously, but this time with a cancellation handler implemented. In this example, `clearInterval()` will run when the `incrementValue` function is canceled.

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## Saving and sharing state

Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.

The following code sample shows an implementation of the previous temperature-streaming function that saves state globally. Note the following about this code:

- `refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second. New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value. It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.

- `streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source. It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.

- Users may call `streamTemperature` from several cells in the Excel UI. Each call reads data from the same `savedTemperatures` variable.

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
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

## Working with ranges of data

Your custom function may accept a range of data as an input parameter, or it may return a range of data. In JavaScript, a range of data is represented as a 2-dimensional array.

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

## Handling errors

When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.

```js
function getComment(x) {
    let url = "https://www.contoso.com/comments/" + x;

    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## Known issues

- Help URLs and parameter descriptions are not yet used by Excel.
- Custom functions are not currently available on Excel for mobile clients.
- Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.
- Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.
- Custom functions in Excel Online may stop working during a session after a period of inactivity. Refresh the browser page (F5) and re-enter a custom function to restore the feature.
- You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows. Close all Excel windows and restart Excel.
- Debugging tools specifically for custom functions may be available in the future. In the meantime, you can debug on Excel Online using F12 developer tools. See more details in [Custom functions best practices](custom-functions-best-practices.md).

## Changelog

- **Nov 7, 2017**: Shipped* the custom functions preview and samples
- **Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later
- **Nov 28, 2017**: Shipped* support for cancellation on asynchronous functions (requires change for streaming functions)
- **May 7, 2018**: Shipped* support for Mac, Excel Online, and synchronous functions running in-process
- **September 20, 2018**: Shipped support for custom functions JavaScript runtime. For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).

\* to the Office Insiders Channel

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](excel-tutorial-custom-functions.md)