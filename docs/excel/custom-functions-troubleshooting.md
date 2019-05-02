---
ms.date: 05/01/2019
description: Troubleshoot common problems in Excel custom functions.
title: Troubleshoot custom functions (preview)
localization_priority: Priority
---
# Troubleshoot custom functions

When developing custom functions, you may encounter errors in the product while creating and testing your functions.

To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages). Also, check for common mistakes such as not [verifying SSL certificates](#my-add-in-wont-load-verify-certificates) properly, [leaving promises unresolved](#ensure-promises-return), and forgetting to [associate your functions](#my-functions-wont-load-associate-functions).

## Enable runtime logging

If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in). Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues. The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.  For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).  

### Check for Excel error messages

Excel has a number of built-in error messages which are returned to a cell if there is calculation error. Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.

Generally, these errors correspond to the errors you might already be familiar with in Excel. The are only a few exceptions specific to custom functions, listed here:

- A `#NAME` error generally means there has been an issue registering your functions.
- A `#VALUE` error typically indicates an error in the functions' script file.
- A `#N/A` error is also maybe a sign that that function while registered could not be run. This is typically due to a missing `CustomFunctions.associate` command.

## Common issues

### My add-in won't load: verify certificates

If your add-in fails to install, verify that the SSL certificates are configured correctly for the web server that's hosting your add-in. Typically if there is a problem with SSL certificates, you will see an error message in Excel warning you that your add-in could not be installed properly. For more information, see [Installing the self-signed certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

### My functions won't load: associate functions

In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md). This is done by using the `CustomFunctions.associate()` method. Typically this method call is made after each function or at the end of the script file. If a custom function is not associated, it will not work.

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

For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).

### Can't open add-in from localhost: use a local loopback exception

If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception. For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/en-us/help/4490419/local-loopback-exemption-does-not-work).

### Ensure promises return

When Excel is waiting for a custom function to complete, it displays #BUSY! in the cell. If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!. Check your functions to make sure that any promises are properly returning a result to a cell.

## Reporting Feedback

If you are encountering issues that aren't documented here, let us know. There are two ways to report issues.

### In Excel on Windows or Mac

If using Excel for Windows or Mac, you can report feedback to the Office extensibility team directly from Excel. To do this, select **File -> Feedback -> Send a Frown**. Sending a frown will provide the necessary logs to understand the issue you are hitting.

### In Github

Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
