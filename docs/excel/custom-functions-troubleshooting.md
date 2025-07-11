---
ms.date: 05/20/2025
description: Troubleshoot common problems with Excel custom functions.
title: Troubleshoot custom functions
ms.topic: troubleshooting
ms.localizationpriority: medium
---
# Troubleshoot custom functions

When developing custom functions, you may encounter errors in the product while creating and testing your functions.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages). Also, check for common mistakes such as [leaving promises unresolved](#ensure-promises-return).

## Debugging custom functions

To debug custom functions add-ins that use a [shared runtime](../testing/runtimes.md#shared-runtime), see [Overview of debugging Office Add-ins](../testing/debug-add-ins-overview.md).

To debug custom functions add-ins that don't use a shared runtime, see [Custom functions debugging](custom-functions-debugging.md).

## Enable runtime logging

If you're testing your add-in in Office on Windows, you should [enable runtime logging](../testing/runtime-logging.md). Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues. The statements cover a variety of errors, including errors pertaining to your add-in's manifest file, runtime conditions, or installation of your custom functions. For more information about runtime logging, see [Debug your add-in with runtime logging](../testing/runtime-logging.md).

### Check for Excel error messages

Excel has a number of built-in error messages which are returned to a cell if there is calculation error. Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.

Generally, these errors correspond to the errors you might already be familiar with in Excel. The are only a few exceptions specific to custom functions, listed here:

- A `#NAME` error generally means there has been an issue registering your functions.
- A `#N/A` error is also maybe a sign that that function while registered could not be run. This is typically due to a missing `CustomFunctions.associate` command.
- A `#VALUE` error typically indicates an error in the functions' script file.
- A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.

## Clear the Office cache

Information about custom functions is cached by Office. Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear. You can fix this by clearing the Office cache. For more information, see [Clear the Office cache](../testing/clear-cache.md).

### Clear the custom functions cache when your add-in runs

There may be times when you need to clear the custom functions cache for an add-in deployed to your end users, so that add-in updates and custom functions setting changes are incorporated at the same time. Without triggering a custom functions cache clear, changes to the **functions.json** and **functions.js** files may take up to 24 hours to reach your end users, while changes to **taskpane.html** reach end users more quickly.

> [!NOTE]
> Once this setting is turned on for a document, it takes effect the next time the document is opened with the add-in. It doesn't apply immediately after the function is called.

To ensure that the custom functions cache is cleared by your add-in, add the following code to your **functions.js** file, and then call the `setForceRefreshOn` method in your `Office.onReady` call or other add-in initialization logic.

> [!IMPORTANT]
> This process for clearing the custom functions cache is only supported for custom functions add-ins that use a [shared runtime](../testing/runtimes.md#shared-runtime).

```javascript
// To enable custom functions cache clearing, add this method to your functions.js 
// file, and then call the `setForceRefreshOn` method in your `Office.onReady` call.
function setForceRefreshOn() {  
    Office.context.document.settings.set(  
        'Office.ForceRefreshCustomFunctionsCache',
        true  
    );  
    Office.context.document.settings.saveAsync();  
}
```

> [!TIP]
> Frequently refreshing the custom functions cache can impact performance, so clearing the custom functions cache with `setForceRefreshOn` should only be used during add-in development or to resolve bugs. Once a custom functions add-in is stabilized, stop forcing cache refreshes.

To disable the custom functions cache clear in your add-in, set `Office.ForceRefreshCustomFunctionsCache` to `false` and call the method in your `Office.onReady` call. The following code sample shows an example with a `setForceRefreshOff` method.

```javascript
// To disable custom functions cache clearing, add this method to your functions.js 
// file, and then call the `setForceRefreshOff` method in your `Office.onReady` call.
function setForceRefreshOff() {  
    Office.context.document.settings.set(  
        'Office.ForceRefreshCustomFunctionsCache',
        false  
    );  
    Office.context.document.settings.saveAsync();  
}
```

## Common problems and solutions

### Can't open add-in from localhost: Use a local loopback exemption

If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exemption. For details on how to do this, see [this Microsoft support article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

### Runtime logging reports "TypeError: Network request failed" on Excel on Windows

If you see the error "TypeError: Network request failed" in your [runtime log](custom-functions-troubleshooting.md#enable-runtime-logging) while making calls to your localhost server, you'll need to enable a local loopback exception. For details on how to do this, see *Option #2* in [this Microsoft support article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

### Ensure promises return

When Excel is waiting for a custom function to complete, it displays #BUSY! in the cell. If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing `#BUSY!`. Check your functions to make sure that any promises are properly returning a result to a cell.

### Error: The dev server is already running on port 3000

Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses). You can stop the dev server by running `npm stop` or by closing the Node.js window. In some cases, it can take a few minutes for the dev server to stop running.

### My functions won't load: associate functions

In cases where your JSON has not been registered and you have authored your own JSON metadata, you may see a `#VALUE!` error or receive a notification that your add-in cannot be loaded. This usually means you need to associate each custom function with its `id` property specified in the [JSON metadata file](custom-functions-json.md). This is done by using the `CustomFunctions.associate()` function. Typically this function call is made after each function or at the end of the script file. If a custom function is not associated, it will not work.

The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

For more information on this process, see [Associate function names with JSON metadata](../excel/custom-functions-json.md#associate-function-names-with-json-metadata).

## Known issues

Known issues are tracked and reported in the [Excel Custom Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## Reporting feedback

If you are encountering issues that aren't documented here, let us know. There are two ways to report issues.

### In Excel on Windows or on Mac

If using Excel on Windows or on Mac, you can report feedback to the Office extensibility team directly from Excel. To do this, select **File** > **Feedback** > **Send a Frown**. Sending a frown will provide the necessary logs to understand the issue you are hitting.

### In Github

Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## Next steps

Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).

## See also

- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Create custom functions in Excel](custom-functions-overview.md)
- [Custom functions debugging](custom-functions-debugging.md)