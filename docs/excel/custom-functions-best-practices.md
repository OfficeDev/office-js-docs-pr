---
ms.date: 09/20/2018
description: Learn best practices and recommended patterns for Excel custom functions.
title: Custom functions best practices
---

# Custom functions best practices

This article describes best practices for developing custom functions in Excel.

## Error handling

When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
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

## Debugging
Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. Then you can debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md). Use `console.log` statements within your custom functions code to send output to the console in real time.

If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.

If you are testing your add-in in Office 2016 desktop you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions. 


## Mapping names

By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel. However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel. This is helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names. `CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.  
  
The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI. When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

The following code sample defines a two key-value pairs. The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI. When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run. When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
