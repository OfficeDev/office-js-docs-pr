---
title: Error handling with the application-specific JavaScript APIs
description: Learn about Excel, Word, PowerPoint, and other application-specific JavaScript API error handling logic to account for runtime errors.
ms.date: 09/27/2022
ms.localizationpriority: medium
---


# Error handling with the application-specific JavaScript APIs

When you build an add-in using the [application-specific Office JavaScript APIs](../develop/application-specific-api-model.md), be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the APIs.

## Best practices

In our [code samples](https://github.com/OfficeDev/Office-Add-in-samples) and [Script Lab](../overview/explore-with-script-lab.md) snippets, you'll notice that every call to `Excel.run`, `PowerPoint.run`, or `Word.run` is accompanied by a `catch` statement to catch any errors. We recommend that you use the same pattern when you build an add-in using the application-specific APIs.

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add your Excel JavaScript API calls here.

      // Await the completion of context.sync() before continuing.
    await context.sync();
    console.log("Finished!");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

```

## API errors

When an Office JavaScript API request fails to run successfully, the API returns an error object that contains the following properties.

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

- **message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.

- **debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.

> [!NOTE]
> If you use `console.log()` to print error messages to the console, those messages are only visible on the server. End users do not see those error messages in the add-in task pane or anywhere in the Office application. To report errors to the user, see [Error notifications](#error-notifications).

## Error Codes and messages

The following tables list the errors that application-specific APIs may return.

> [!NOTE]
> The preceding table lists error messages you may encounter while using the application-specific APIs. If you are working with the Common API, see [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md) to learn about relevant error messages.

|Error code | Error message | Notes |
|:----------|:--------------|:------|
|`AccessDenied` |You cannot perform the requested operation.|*None.* |
|`ActivityLimitReached`|Activity limit has been reached.|*None.* |
|`ApiNotAvailable`|The requested API is not available.|*None.* |
|`ApiNotFound`|The API you are trying to use could not be found. It may be available in a newer version of Excel. See the [Excel JavaScript API requirement sets](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) article for more information.|*None.* |
|`BadPassword`|The password you supplied is incorrect.|*None.* |
|`Conflict`|Request could not be processed because of a conflict.|*None.* |
|`ContentLengthRequired`|A `Content-length` HTTP header is missing.|*None.* |
|`GeneralException`|There was an internal error while processing the request.|*None.* |
|`InsertDeleteConflict`|The insert or delete operation attempted resulted in a conflict.|*None.* |
|`InvalidArgument` |The argument is invalid or missing or has an incorrect format.|*None.* |
|`InvalidBinding` |This object binding is no longer valid due to previous updates.|*None.* |
|`InvalidOperation`|The operation attempted is invalid on the object.|*None.* |
|`InvalidOperationInCellEditMode`|The operation isn't available while Excel is in Edit cell mode. Exit Edit mode by using the **Enter** or **Tab** keys, or by selecting another cell, and then try again.|*None.* |
|`InvalidReference`|This reference is not valid for the current operation.|*None.* |
|`InvalidRequest`  |Cannot process the request.|*None.* |
|`InvalidSelection`|The current selection is invalid for this operation.|*None.* |
|`ItemAlreadyExists`|The resource being created already exists.|*None.* |
|`ItemNotFound` |The requested resource doesn't exist.|*None.* |
|`MemoryLimitReached`|The memory limit has been reached. Your action could not be completed.|*None.* |
|`NotImplemented`|The requested feature isn't implemented.| This could mean the API is in preview or only supported on a particular platform (such as online-only). See [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets) for more information.|
|`RequestAborted`|The request was aborted during run time.|*None.* |
|`RequestPayloadSizeLimitExceeded`|The request payload size has exceeded the limit. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.| This error only occurs in Office on the web.|
|`ResponsePayloadSizeLimitExceeded`|The response payload size has exceeded the limit. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.|  This error only occurs in Office on the web.|
|`ServiceNotAvailable`|The service is unavailable.|*None.* |
|`Unauthenticated` |Required authentication information is either missing or invalid.|*None.* |
|`UnsupportedFeature`|The operation failed because the source worksheet contains one or more unsupported features.|*None.* |
|`UnsupportedOperation`|The operation being attempted is not supported.|*None.* |

### Excel-specific error codes and messages

|Error code | Error message | Notes |
|:----------|:--------------|:------|
|`EmptyChartSeries`|The attempted operation failed because the chart series is empty.|*None.* |
|`FilteredRangeConflict`|The attempted operation causes a conflict with a filtered range.|*None.* |
|`FormulaLengthExceedsLimit`|The bytecode of the applied formula exceeds the maximum length limit. For Office on 32-bit machines, the bytecode length limit is 16384 characters. On 64-bit machines, the bytecode length limit is 32768 characters.| This error occurs in both Excel on the web and on desktop.|
|`GeneralException`|*Various.*|The data types APIs return `GeneralException` errors with dynamic error messages. These messages reference the cell that is the source of the error, and the problem that is causing the error, such as: "Cell A1 is missing the required property `type`."|
|`InactiveWorkbook`|The operation failed because multiple workbooks are open and the workbook being called by this API has lost focus.|*None.* |
|`MergedRangeConflict`|Cannot complete the operation. A table can't overlap with another table, a PivotTable report, query results, merged cells, or an XML Map.|*None.* |
|`NonBlankCellOffSheet`|Microsoft Excel can't insert new cells because it would push non-empty cells off the end of the worksheet. These non-empty cells might appear empty but have blank values, some formatting, or a formula. Delete enough rows or columns to make room for what you want to insert and then try again.|*None.* |
|`OperationCellsExceedLimit`|The attempted operation affects more than the limit of 33554000 cells.| If the `TableColumnCollection.add API` triggers this error, confirm that there is no unintentional data within the worksheet but outside of the table. In particular, check for data in the right-most columns of the worksheet. Remove the unintended data to resolve this error. One way to verify how many cells that an operation processes is to run the following calculation: `(number of table rows) x (16383 - (number of table columns))`. The number 16383 is the maximum number of columns that Excel supports. <br><br>This error only occurs in Excel on the web. |
|`PivotTableRangeConflict`|The attempted operation causes a conflict with a PivotTable range.|*None.* |
|`RangeExceedsLimit`|The cell count in the range has exceeded the maximum supported number. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.|*None.* |
|`RefreshWorkbookLinksBlocked`|The operation failed because the user hasn't granted permission to refresh external workbook links.|*None.* |
|`UnsupportedSheet`|This sheet type does not support this operation, since it is a Macro or Chart sheet.|*None.* |

## Error notifications

How you report errors to users depends on the UI system you are using. If you're using React as the UI system, use the Fluent UI components and design elements. Pick an appropriate control from this [Fluent UI page](https://developer.microsoft.com/fluentui#/controls/web). We recommend that error messages be conveyed with a message bar, dialog, or modal. If the error is in the user's input, display the error in bold red near the input control. The sample [Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) uses a MessageBar element and modifies it to take account of the personality menu in an add-in task pane.

If you're not using React for the UI, consider using the older Fabric UI components implemented directly in HTML and JavaScript. Some example templates are in the [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) repository. Take a look especially in the dialog and navigation subfolders. The sample [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) uses a message banner.

## See also

- [OfficeExtension.Error object (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md)
