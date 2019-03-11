---
ms.date: 03/06/2019
description: Troubleshoot common problems in Excel custom functions.
title: Troubleshoot custom functions (preview)
localization_priority: Priority
---
# Troubleshoot custom functions

When developing custom functions, you may encounter errors in the product while creating and testing your functions.

To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages). Also, check for common mistakes such as not [verifying ssl certificates](#verify-ssl-certificates) properly, [leaving promises unresolved](#ensure-promises-return), and forgetting to [associate your functions](#associate-your-functions).

## Enable runtime logging

If you are testing your add-in in Office on Windows, you should [enable runtime logging](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in) to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions. Runtime logging writes `console.log` statements to a log file to help you uncover issues. For more information about runtime logging, see [Use runtime logging to debug your add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).  

### Check for Excel error messages

Excel has a number of built-in error messages which are returned to a cell if there is calculation error. Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#GETTING_DATA`.

## Common issues

### Verify SSL certificates

If your add-in fails to register, [verify that the SSL certificates are configured correctly](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application. Typically if you forget to do this step, you will see an error message in Excel warning you that custom functions could not be installed properly. For more information on this verification, see [Adding self-signed certificates as trusted root certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

### Associate your functions

In your custom functions' script file, you need to both write the function code and code to associate the function's name to the corresponding JSON id. Typically this declaration is made after each function or at the end of the script file. It's common to forget to include this and can be a reason why your functions don't work.

The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

For more information on this process, see [Associating function names with json metadata](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).

### Ensure promises return

In addition to the ordinary reasons that a cell might report a #GETTING_DATA error, custom functions also will report `#GETTING_DATA` if a promise does not return a result. Check your functions to make sure that any promises are properly returning a result to a cell.

## Reporting Feedback

If you are encountering errors that aren't documented here, let us know. The custom functions and documentation teams have several procedures in place for reporting issues.

### In Excel on Windows or Mac

If using Excel for Windows or Mac, you can report feedback to the Excel Custom Functions team directly in the UI of Excel itself. To do this, select **File | Feedback | Send a Frown**. Sending a frown will provide the necessary logs to understand the issue you are hitting.

### In Github

Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
