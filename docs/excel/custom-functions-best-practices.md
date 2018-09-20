---
ms.date: 09/20/2018
description: Learn best practices and recommended patterns for Excel custom functions.
title: Custom Functions' best practices
---

# Custom Functions' best practices

This article describes some recommended patterns and solutions to common use cases with Excel custom functions.

## Error handling

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

## Error logging

You can use runtime logging to debug your custom function's XML manifest file or to look for errors in your custom functions in real time via console.log statements. Runtime logging is only available for Office 2016 desktop currently.

For full instructions on how to use runtime logging, [read this article](../testing/troubleshoot-manifest.md).

## Debugging

At present, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser. Additional debugging tools for custom functions may be available in the future.

## Mapping names

Custom functions are typically declared entirely in uppercase letters, although you can change this by using the  `CustomFunctionsMappings` object. The key-value pairs you specify in `CustomFunctionsMappings` correspond to the function name you call in Excel (such as `=ADD42`) and the new alternate name you would like to use for this function in Excel. Use of `CustomFunctionsMapping` is not required, but can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with the uppercase letters in these functions.
  
You can declare individual functions, as shown below:  

```js
function ADD42(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42 //effectively renames the add-in when invoked in Excel, so you will now call =plusFortyTwo()
}
```

However, you can declare multiple mappings at the same time, as shown in the example below.  

```js
//assume that COUNTDOGS and COUNTCATS exist
  
CustomFunctionsMappings = {
    "countdogs" : COUNTDOGS,  
    "meow" : COUNTCATS
}
 ```