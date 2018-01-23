---
title: Create custom functions in Excel (Preview)
description: ''
ms.date: 01/23/2018
---

# Create custom functions in Excel (Preview)

Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in. Users can then access custom functions like any other native function in Excel (like =SUM()). This article explains how to create custom functions in Excel.

Here's what custom functions look like in Excel:

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Here’s the code for a sample custom function that adds 42 to a pair of numbers.

```js
function add42 (a, b) {
    return a + b + 42;
}
```

Custom functions are now available in preview. Follow these steps to try them:

1.  Join the [Office Insider](https://products.office.com/en-us/office-insider) program to install the version of Excel 2016 that's required for custom functions on your computer (version 16.8711 or later). You must choose the "Insider" channel for the custom functions preview to work.
2.  Clone the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repo and follow the instructions in *README.md* to start the add-in in Excel.
3.  Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.
4.  If you have questions, ask them on Stack Overflow with the [office-js](https://stackoverflow.com/questions/tagged/office-js) tag.

See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.

## Learn the basics


In the cloned sample repo, you’ll see the following files:

-   *customfunctions.js*, which contains:

    -   The custom function code to add to Excel.
    -   The registration code to connect your custom function to Excel. Registration makes your custom functions appear in the list of available functions displayed when users type in cells.
-   *customfunctions.html*, which provides a &lt;Script&gt; reference to *customfunctions.js*. This file does not display UI in Excel.
-   *manifest.xml*, which tells Excel the location of your HTML and JS files needed to run custom functions.

### JavaScript file (*customfunctions.js*)

The following code in customfunctions.js declares the custom function `add42`, and then registers the function in Excel.

```js
function add42 (a, b) {
    return a + b + 42;
}

Excel.Script.customFunctions["CONTOSO"]["ADD42"] = {
    call: add42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [{
        name: "num 1",
        description: "The first number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    {
        name: "num 2",
        description: "The second number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    }],
    options:{ batch: false, stream: false }
};

Excel.run(function(ctx) {
    ctx.workbook.customFunctions.addAll();
});
```

**Registration** of the custom function uses the `Excel.Script.customFunctions["CONTOSO"]["ADD42"]` code block. You need the following parameters to register the function in Excel:

-   Prefix and function name: The first value in `Excel.Script.customFunctions` is the prefix (in this case, CONTOSO is the prefix). The second value in `Excel.Script.customFunctions` is the function name (in this case ADD42 is the function name). In Excel, the prefix and the function name are separated using a period: to use your custom function, combine the function's prefix (CONTOSO) with the function's name (ADD42) and enter `=CONTOSO.ADD42` into a cell. By convention, prefixes and function names use upper case letters. The prefix is intended to be used as an identifier for your add-in.
-   `call`: Defines the JavaScript function to call (for example, `add42`). The name of the JavaScript function does not need to match the name that you register in Excel.
-   `description`: The description appears in the autocomplete menu in Excel.
-   `helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.
-   `result`: Defines the type of information returned by the function to Excel.

    -   `resultType`: Your function can return either a `"string"` or a `"number"` (also used for dates and currencies). For more information see [
Custom Function Enumerations](https://dev.office.com/reference/add-ins/excel/customfunctionsenumerations).
    -   `resultDimensionality`: Your function can return either a single (`"scalar"`) value or a `"matrix"` of values. When returning a matrix of values, your function returns an array, where each array element is another array that represents a row of values. For more information, see [Custom Function Enumerations](https://dev.office.com/reference/add-ins/excel/customfunctionsenumerations). The following example returns a 3-row, 2-column matrix of values from a custom function.

        ```js
        return [["first","row"],["second","row"],["third","row"]];
        ```

-   Your custom function may take arguments as input. The arguments passed to your custom function are specified in the *parameters* property. The order of the parameters in the definition must match the order in the JavaScript function. For each parameter, define these properties:

    -   `name`: The string displayed in Excel to represent the parameter.
    -   `description`: The string displayed for more information about the parameter.
    -   `valueType`: A `"number"` or `"string"`, similar to the resultType property described earlier.
    -   `valueDimensionality`: A `"scalar"` value or `"matrix"` of values, similar to the resultDimensionality property described previously. Matrix-type parameters allow the user to select ranges larger than a single cell.

-   `options`: enables special types of custom functions that are described in more detail later in this article.

To complete registration of all functions defined using `Excel.Script.customFunctions`, ensure you call `CustomFunctions.addAll()`.

After registration, custom functions are available in all workbooks (not only the one where the add-in ran initially) for a user. The functions are displayed in the autocomplete menu when the user starts typing it. During development and testing, you can manually clear your computer's cache of registration metadata by deleting the folder `<user>\AppData\Local\Microsoft\Office\16.0\Wef\CustomFunctions`.


### Manifest file (*manifest.xml*)

The following example in manifest.xml allows Excel to locate the code for your functions.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">

    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="scriptURL" />
                        <!— Required. The Developer Preview does not use the Script element.-->
                    </Script>
                    <Page>
                        <SourceLocation resid="pageURL"/>
                    </Page>
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>

    <Resources>
        <bt:Urls>
            <bt:Url id="scriptURL" DefaultValue="https://www.contoso.com/addin/customfunctions.js" />
            <bt:Url id="pageURL" DefaultValue="https://www.contoso.com/addin/customfunctions.html" />
        </bt:Urls>
    </Resources>

</VersionOverrides>

```

The previous code specifies:

-   A `<Script>` element, which is required but not used in the Developer Preview.
-   A `<Page>` element, which links to the HTML page of your add-in. The HTML page includes a &lt;Script&gt; reference to the JavaScript file (*customfunctions.js*) that contains the custom function and registration code. The HTML page is a hidden page and is never displayed in the UI.

## Asynchronous functions

If your custom function retrieves data from the web, you need to make an asynchronous call to fetch it. When calling external web services, your custom function must:

1.   Return a JavaScript Promise to Excel.
2.   Make the http request to call the external service.
3.   Resolve the promise using the `setResult` callback. `setResult` sends the value to Excel.

The following code shows an example of a custom function that retrieves the temperature of a thermometer.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult, setError){
        sendWebRequestExample(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## Streamed functions

Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations. For example, the `incrementValue` custom function in the following code adds a number to the result every second, and Excel displays each new value automatically using the `setResult` callback. To see the registration code used with `incrementValue`, read the *customfunctions.js* file.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function. It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell. In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the parameter `stream` to `true`.

## Cancellation

You can cancel streamed functions and asynchronous functions. Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load. Excel cancels function calls in the following situations:
- The user edits or deletes a cell that references the function.
- One of the arguments (inputs) for the function changes. In this case, a new function call is triggered in addition to the cancelation.
- The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.

The following code shows the previous example with cancellation implemented. In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## Saving state

Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users enter multiple instances of the same custom function, and they need to share data with each other. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.

The following code shows an implementation of the previous temperature-streaming function that saves state using the `savedTemperatures` variable. The code demonstrates the following concepts:

-   **Saving data.** `refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second. New temperatures are saved in the savedTemperatures variable.

-   **Using saved data.** `streamTemperature` updates the temperature values displayed in the Excel UI every second. Temperatures are read from `savedTemperature`, and then sent to the Excel UI using `setResult`. Users may call `streamTemperature` from several cells in the Excel UI. Each call to `streamTemperature` will read data from `savedTemperatures`.

> In this case, we register `streamTemperature` as the custom function in Excel.

```js
var savedTemperatures{};

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequestExample(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## Working with ranges of data

Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.

For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel. The following function takes the parameter `temperatures`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.

```js
function secondHighestTemp(temperatures){ 
     var highest = -273, secondHighest = -273;
     for(var i = 0; i < temperatures.length; i++){
         for(var j = 0; j < temperatures[i].length; j++){
             if(temperatures[i][j] <= highest){
                 secondHighest = highest;
                 highest = temperatures[i][j];
             }
             else if(temperatures[i][j] <= secondHighest){
                 secondHighest = temperatures[i][j];
             }
         }
     }
     return secondHighest;
 }
```

If you create a function that returns a range of data, it's necessary to enter an Array Formula in Excel to see the whole range of values. For more information, see [Guidelines and examples of array formulas](https://support.office.com/en-us/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7).

## Known issues

The following features aren't yet supported in the Developer Preview.

-   Batching, which allows you to aggregate multiple calls to the same function to improve performance.

-   Help URLs and parameter descriptions are not yet used by Excel.

-   Publishing add-ins that use custom functions to AppSource or via Office 365 centralized deployment.

-   Custom functions are not available on Excel on Mac, Excel for iOS, and Excel Online.

-   Currently, add-ins rely on a hidden browser process to run custom functions. In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory. Additionally, the HTML page referenced by the &lt;Page&gt; element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly. To prepare for this change, ensure your custom functions do not use the webpage DOM.

## Changelog

- **Nov 7, 2017**: Shipped the custom functions preview and samples
- **Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later
- **Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)
