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

## Debugging
Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md). Use `console.log` statements within your custom functions code to send output to the console in real time.

If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.

If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.


## Mapping names

When using the `CustomFunctionMappings` object, there are several best practices to keep in mind.

* Function names and ids should match in the functions JSON file.

    Using the same name and id for a function keeps your code simple. Using this pattern, you will not have to track two separate ways to refer to the same function.  

* Function names in JSON should be listed in uppercase letters.

    Using uppercase names is standard for the functions which are already built into Excel. Using this established pattern, your custom function can seamlessly integrate into Excel's existing user experience.

* Declare only one instance of `CustomFunctionMappings` in your JavaScript code in order to prevent overwriting functions with new mappings.

    If a function is mapped in `CustomFunctionMappings`, this can be overwritten by another declaration of `CustomFunctionMappings` later in your code. As shown in the following example: 

    ```js
    addNine(x) {
        return x + 9;
    }

    CustomFunctionMappings = {
        ADDNINE = addNine;
    }

    addTen(x) {
       return x + 10;
    }

    CustomFunctionMappings = {
        ADDNINE = addTen; //Now when Excel users call ADDNINE, they will get addTen
    }

 ## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Custom functions metadata](custom-functions-json.md)
- [Runtime for Excel custom functions](custom-functions-runtime.md)
