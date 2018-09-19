---
ms.date: 09/18/2018
description: Create your own custom function add-in in Excel using JavaScript. 
title: Create Custom Functions in Excel (Preview)
---

# Create custom functions in Excel (Preview)

Custom functions (similar to user-defined functions, or UDFs), enable developers to add any JavaScript function to Excel using an add-in. Users can then access custom functions like any other native function in Excel (such as `=SUM()`). This article explains how to create custom functions in Excel.

The following illustration shows you how an end user would insert a custom function into a cell. The function that adds 42 to a pair of numbers.

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Here’s the code for the same custom function.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online. Follow these steps to try them:

1. Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program. (Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)
2. Create an Excel Custom Functions add-in project using [Yo Office](https://github.com/OfficeDev/generator-office), and follow the instructions in the [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to start the add-in in Excel, make changes in the code, and debug.
3. Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.

See the **Known Issues** section at the end of this article, which includes current limitations of custom functions and will be updated over time.

## Learn the basics

In the cloned sample repo, you’ll see the following files:

- **./src/customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).
- **./config/customfunctions.json**, which contains the registration JSON that tells Excel about your custom function. Registration makes your custom functions appear in the list of available functions displayed when a user types in a cell.
- **./index.html**, which provides a &lt;Script&gt; reference to the JS file.
- **./manifest.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; it also specifies a namespace for all the custom functions that are installed with the add-in.

### JSON file (./config/customfunctions.json)

The following code in **customfunctions.json** specifies the metadata for the same `ADD42` function. This metadata includes details on the function's name, description, returned value, parameters, and more. Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](custom-functions-json.md).

> [!NOTE]
> To call a function in Excel, the name listed in this JSON file will be appended to the namespace specified in your XML manifest file. To call this function, you'd call `=CONTOSO.ADD42`. The namespace is intended to be used as an identifier for your company or the add-in. A namespace and the name of a function are separated by a period.

```js
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [ //array can contain multiple functions
        {
            "name": "ADD42", //names your function, recommended that you use uppercase letters
            "description":  "adds 42 to the input numbers", //appears in the autocomplete menu in Excel
            "helpUrl": "http://dev.office.com", //Excel opens a taskpane and displays the webpage listed here when a user requests help with a function
            "result": { //specifies the type information returned by the function to Excel
                "type": "number", //can be string, number, boolean
                "dimensionality": "scalar" //can be scalar or matrix (two dimensional array of values specified by type above)
            },
            "parameters": [ //parameters must be specified in order
                {
                    "name": "number 1", //appears in Excel IntelliSense
                    "description": "the first number to be added", //appears in Excel IntelliSense
                    "type": "number", //should mirror the type of result 
                    "dimensionality": "scalar" //should mirror the dimensionality of result
                },
                {
                    "name": "number 2", //appears in Excel IntelliSense. Remember that parameters are in order.
                    "description": "the second number to be added", //appears in Excel IntelliSense
                    "type": "number", //should mirror type of result
                    "dimensionality": "scalar" //should mirror the dimensionality of result
                }
            ],
        }
    ]
}
```

> [!NOTE]
> The custom functions are registered when a user runs the add-in for the first time. After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)

Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.

### Manifest file (./manifest.xml)

The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions. This allows you to change the locations of your JSON, JavaScript, and HTML files which make up your custom function.  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
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

## Initializing custom functions

Your code must initialize the custom functions feature before using it. You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js). During the preview of custom functions, you have your choice of two syntaxes for intializing. The HTML file in the repo uses the following syntax:

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

You can also use the following syntax:

```js
Office.Preview.StartCustomFunctions();
```

## Handling errors

Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](./excel-add-ins-error-handling.md). Generally, you will use `.catch` to handle errors. The code below gives an example of `.catch`.

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
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

## Asynchronous functions

If your custom function retrieves data from the web, it must be asynchronous with respect to Excel. Asynchronous functions must:

1. Return a JavaScript Promise to Excel.
2. Resolve the Promise with the final value using the callback function.

The following code shows an example of an asynchronous custom function that retrieves the temperature of a thermometer. Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result. Users can interact normally with the rest of the spreadsheet while they wait for the result.

## Streamed functions

An asynchronous function can be streamed. Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations. The following example is a custom function that adds a number to the result every second. Note the following about this code:

- Excel displays each new value automatically using the `setResult` callback.
- The final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function. It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.
- In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## Cancellation

You can cancel streamed functions and asynchronous functions. Canceling your function calls is important to reduce their bandwidth consumption, working memory, and CPU load. Excel cancels function calls in the following situations:

- The user edits or deletes a cell that references the function.
- One of the arguments (inputs) for the function changes. In this case, a new function call is triggered in addition to the cancellation.
- The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancellation.

You *must* implement a cancellation handler for every streaming function. Asynchronous, non-streaming functions may or may not be cancelable; it's up to you. Synchronous functions cannot be canceled.

To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.

The following code shows the previous example with cancellation implemented. In the code, the `handler` object contains an `onCanceled` function must be defined for each cancelable custom function.

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

Asynchronous custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.

The following code shows an implementation of the previous temperature-streaming function that saves state globally. Note the following about this code:

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

> [!NOTE]
> Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.

## Working with ranges of data

Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.

For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type. Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.

```js
function secondHighest(values){
     var highest = values[0][0], secondHighest = values[0][0];
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

As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).

## Known issues

- Help URLs and parameter descriptions are not yet used by Excel.
- Custom functions are not currently available on Excel for mobile clients.
- Currently, add-ins rely on a hidden browser process to run asynchronous custom functions. In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory. Additionally, the HTML page referenced by the `<Page>` element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly. To prepare for this change, ensure your custom functions do not use the web page DOM. The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) using GET or POST.
- Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.
- Debugging is only enabled for asynchronous functions on Excel for Windows.
- Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.
- Custom functions in Excel Online may stop working during a session after a period of inactivity. Refresh the browser page (F5) and re-enter a custom function to restore the feature.

## Changelog

- **Nov 7, 2017**: Shipped* the custom functions preview and samples
- **Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later
- **Nov 28, 2017**: Shipped* support for cancellation on asynchronous functions (requires change for streaming functions)
- **May 7, 2018**: Shipped* support for Mac, Excel Online, and synchronous functions running in-process