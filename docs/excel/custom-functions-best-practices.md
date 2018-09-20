---
ms.date: 09/20/2018
description: Learn best practices and recommended patterns for Excel custom functions.
title: Custom functions best practices
---

# Custom functions best practices

This article describes best practices for developing custom functions in Excel.

## Error handling

When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; 
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

## Error logging

You can enable error logging for your custom functions add-in in multiple ways, such as: 

- [Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file. 

- Use `console.log` statements within your custom functions code to send output to the console in real time.

> [!NOTE]
> Runtime logging is currently available only for Office 2016 desktop.

## Debugging

Currently, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser. Additional debugging tools for custom functions may be available in the future.

## Mapping names

By default, the name of a custom function in your JavaScript file should be declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel. However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel. Although you're not required to use `CustomFunctionsMapping`, it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.
  
The following code sample defines a single key-value pair that maps the JavaScript function name `ADD42` to the `plusFortyTwo` function name in the Excel UI. When the end user chooses the `plusFortyTwo` function in Excel, the `ADD42` JavaScript function will run.

```js
function ADD42(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

The following code sample defines a two key-value pairs. The first pair maps the JavaScript function name `ADD50` to the `plusFifty` function name in the Excel UI, and the second pair maps the JavaScript function name `ADD100` to the `plusOneHundred` function name in the Excel UI. When the end user chooses the `plusFifty` function in Excel, the `ADD50` JavaScript function will run. When the end user chooses the `plusOneHundred` function in Excel, the `ADD100` JavaScript function will run.

```js
function ADD50(num) {
    return num + 50;  
} 

function ADD100(num) {
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