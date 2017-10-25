# Create custom functions in Excel (Preview)

Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in. Users can then access custom functions like any other native function in Excel (like =SUM()). This article explains how to create custom functions in Excel.

The following illustration shows you how custom functions work in the Excel UI.

<img src="../../images/custom-function.gif" width="579" height="383" />

Here’s the code for a sample custom function that adds 42 to a pair of numbers.

```js
function add42 (a, b) {
    return a + b + 42;
}
```

Custom functions are now available in preview. Follow these steps to try them:

1.  Install Office 2016 for Windows and join the [Office Insider](https://products.office.com/en-us/office-insider) program.
2.  Clone the Excel-Custom-Functions repo and follow the instructions in the README.md to start the add-in in Excel.
2.  Clone the *Excel-Custom-Functions* repo and follow the instructions in *README.md* to start the add-in in Excel.
3.  Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.

See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.

## Learn the basics


In the cloned sample repo, you’ll see the following files:

-   *customfunctions.js*, which contains:

    -   The custom function code to add to Excel.
    -   The registration code to connect your custom function to Excel. Registration makes your custom functions appear in the list of available functions displayed when users type in cells.
-   *customfunctions.html*, which provides a &lt;Script&gt; reference to customfunctions.js. This file does not display UI in Excel.
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
    parameters: [
    {
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
    }
    ],
    options:{ batch: false, stream: false }
};

Excel.run(function(ctx) {
    ctx.workbook.customFunctions.addAll();
});
```

**Registration** of the custom function takes place within the `Excel.Script.customFunctions["CONTOSO"]["ADD42"]` code block. You need the following parameters to register the function in Excel:

-   Prefix and function name: The first value in `Excel.Script.customFunctions` is the prefix (in this case, CONTOSO is the prefix). The second value in `Excel.Script.customFunctions` is the function name (in this case ADD42 is the function name). The prefix and the function name are separated using a period. To use your custom function, combine the function's prefix (CONTOSO) with the function's name (ADD42) and enter `=CONTOSO.ADD42` into a cell. By convention, prefixes and function names use upper case letters. The prefix is intended to be used as an identifier for your add-in.
-   call: Defines the JavaScript function to call (for example, add42). The name of the JavaScript function does not need to match the name that you register in Excel.
-   description: The description appears in the autocomplete menu in Excel.
-   helpUrl: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.
-   result: The result types of your function can be either a number or string, and either a single value of group of values.

    -   resultType: Your function can return either a number or a string. For more information see &lt;&lt;LINK&gt;&gt;

    -   resultDimensionality: Your function can return either a single value or a matrix of values. When returning a matrix of values, your function returns an array, where each array element is another array of values. For more information, see &lt;&lt;LINK&gt;&gt;. The following example returns a 3-row, 2-column matrix of values from a custom function.

```js
return [[1,1],[2,2],[3,3]];
```

-   Your custom function may take arguments as input. The arguments passed to your custom function are specified in the *parameters* property. If you declare more than 1 parameter, ensure the order of the parameter definition matches your JavaScript function. The parameter property arguments are described as follows:

    -   name: The string displayed in Excel representing the parameter name.
    -   description: A description of the parameter.
    -   valueType: A number or string, that is similar to the resultType described earlier.
    -   valueDimensionality: A single value or matrix of values, that is similar to the resultDimensionality described previously. Matrix-type parameters allow the user to select ranges larger than a single cell.

-   options: { batch: false, stream: false } specifies special types of custom functions that are described in more detail later in this article.

To complete registration of all functions defined using `Excel.Script.Customfunctions`, ensure you call `customFunctions.addAll()`.

After registration, custom functions are available in all workbooks (not only the one where the add-in ran initially) for a user. The functions are displayed in the autocomplete menu when the user starts typing it.

### Manifest file (*manifest.xml*)

The following code in manifest.xml allows Excel to locate the code for your functions.

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

-   A &lt;Script&gt; element, which is required but not used in the Developer Preview.
-   A &lt;Page&gt; element, which links to the HTML page of your add-in. The HTML page includes a &lt;Script&gt; reference to the JavaScript file (customfunctions.js) that contains the custom function and registration code. The HTML page is a hidden page and is never displayed in the UI.

## Asynchronous functions

If your custom function retrieves data from the web, you need to make an asynchronous call to fetch it. When calling external web services, your custom function must:

-   Return a JavaScript Promise to Excel.
-   Make the http request to call the external service (for example, the following code uses sendWebRequestExample).
-   Resolve the promise using the setResult callback. setResult sends the result to Excel.

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

Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations. For example, as shown in the following code, the custom function incrementValue adds a number to the result every second, and Excel displays each new value automatically using the setResult callback. To see the registration code used with incrementValue, see the customfunctions.js file.

```js
function incrementValue(increment, setResult){ 
     var result = 0;
     setInterval(function(){
         result += increment;
         setResult(result);
    }, 1000);
}
```

For streamed functions, the final parameter, setResult, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function. It’s a callback function that’s used to pass data from the function to Excel to update the value of a cell. In order for Excel to pass the setResult function, you must declare support for streaming during your function registration by setting the parameter stream to true.

## Saving state

Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in the global variables. Saved state is useful when users enter multiple instances of the same custom function, and each instance needs to share the same data. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.

The following code shows an implementation of the previous temperature streaming function that saves state using the savedTemperatures variable. The code demonstrates the following concepts:

-   **Saving data.** refreshTemperature is a streamed function that reads the temperature of a particular thermometer every second. New temperatures are saved in the savedTemperatures variable.

-   **Using saved data.** streamTemperature updates the temperature values displayed in the Excel UI every second. Temperatures are read from savedTemperature, and then sent to the Excel UI using setResult. Users may call streamTemperature from several cells in the Excel UI. Each call to streamTemperature will read data from savedTemperatures.

> Note: In this case, we register streamTemperature as the custom function in Excel.

```js
var savedTemperatures{};

function streamTemperature(thermometerID, setResult){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to the Excel UI.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating the Excel UI again.
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

You can use a range of data in your custom function. You can pass a range as a parameter, or you can return a range from a custom function.

For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel. The following function takes the parameter temperatures, which is an Excel.CustomFunctionDimensionality.matrix parameter type.

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

## Known issues

The following features aren't yet supported in the Developer Preview.

-   Batching, which allows you to aggregate multiple calls to the same function to improve performance.

-   Cancelation, which notifies you when a streaming function is no longer required (for example, when users clear a cell). Today, the functions can’t determine when to stop writing new values into the cell.

-   Publishing add-ins to the Office Store or Office 365 centralized deployment that use custom functions.

-   Custom functions are not available on Excel on Mac, Excel for iOS, and Excel Online.

-   Currently, add-ins rely on a hidden browser process to run custom functions. In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory. Additionally, the HTML page referenced by the &lt;Page&gt; element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly. To prepare for this change, ensure your custom functions do not use the webpage DOM.
