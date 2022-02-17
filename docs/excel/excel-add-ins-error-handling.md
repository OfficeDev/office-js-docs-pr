---
title: Error handling with the Excel JavaScript API
description: 'Learn about Excel JavaScript API error handling logic to account for runtime errors.'
ms.date: 02/16/2022
ms.localizationpriority: medium
---


# Error handling with the Excel JavaScript API

When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.

> [!NOTE]
> For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).

## Best practices

In our [code samples](https://github.com/OfficeDev/Office-Add-in-samples) and [Script Lab](../overview/explore-with-script-lab.md) snippets, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add you Excel JavaScript API calls here.

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

When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties.

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

- **message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.

- **debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.

> [!NOTE]
> If you use `console.log()` to print error messages to the console, those messages will only be visible on the server. End users will not see those error messages in the add-in task pane or anywhere in the Office application.

## Error Messages

The following table is a list of errors that the API may return.

|Error code | Error message | Notes |
|:----------|:--------------|:------|
|`AccessDenied` |You cannot perform the requested operation.| |
|`ActivityLimitReached`|Activity limit has been reached.| |
|`ApiNotAvailable`|The requested API is not available.| |
|`ApiNotFound`|The API you are trying to use could not be found. It may be available in a newer version of Excel. See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.| |
|`BadPassword`|The password you supplied is incorrect.| |
|`Conflict`|Request could not be processed because of a conflict.| |
|`ContentLengthRequired`|A `Content-length` HTTP header is missing.| |
|`EmptyChartSeries`|The attempted operation failed because the chart series is empty.| |
|`FilteredRangeConflict`|The attempted operation causes a conflict with a filtered range.| |
|`FormulaLengthExceedsLimit`|The bytecode of the applied formula exceeds the maximum length limit. For Office on 32-bit machines, the bytecode length limit is 16384 characters. On 64-bit machines, the bytecode length limit is 32768 characters.| This error occurs in both Excel on the web and on desktop.|
|`GeneralException`|There was an internal error while processing the request.| |
|`InactiveWorkbook`|The operation failed because multiple workbooks are open and the workbook being called by this API has lost focus.| |
|`InsertDeleteConflict`|The insert or delete operation attempted resulted in a conflict.| |
|`InvalidArgument` |The argument is invalid or missing or has an incorrect format.| |
|`InvalidBinding` |This object binding is no longer valid due to previous updates.| |
|`InvalidOperation`|The operation attempted is invalid on the object.| |
|`InvalidOperationInCellEditMode`|The operation isn't available while Excel is in Edit cell mode. Exit Edit mode by using the **Enter** or **Tab** keys, or by selecting another cell, and then try again.| |
|`InvalidReference`|This reference is not valid for the current operation.| |
|`InvalidRequest`  |Cannot process the request.| |
|`InvalidSelection`|The current selection is invalid for this operation.| |
|`ItemAlreadyExists`|The resource being created already exists.| |
|`ItemNotFound` |The requested resource doesn't exist.| |
|`MemoryLimitReached`|The memory limit has been reached. Your action could not be completed.| |
|`MergedRangeConflict`|Cannot complete the operation. A table can't overlap with another table, a PivotTable report, query results, merged cells, or an XML Map.|
|`NonBlankCellOffSheet`|Microsoft Excel can't insert new cells because it would push non-empty cells off the end of the worksheet. These non-empty cells might appear empty but have blank values, some formatting, or a formula. Delete enough rows or columns to make room for what you want to insert and then try again.| |
|`NotImplemented`|The requested feature isn't implemented.| |
|`OperationCellsExceedLimit`|The attempted operation affects more than the limit of 33554000 cells.| If the `TableColumnCollection.add API` triggers this error, confirm that there is no unintentional data within the worksheet but outside of the table. In particular, check for data in the right-most columns of the worksheet. Remove the unintended data to resolve this error. One way to verify how many cells that an operation processes is to run the following calculation: `(number of table rows) x (16383 - (number of table columns))`. The number 16383 is the maximum number of columns that Excel supports. <br><br>This error only occurs in Excel on the web. |
|`PivotTableRangeConflict`|The attempted operation causes a conflict with a PivotTable range.| |
|`RangeExceedsLimit`|The cell count in the range has exceeded the maximum supported number. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.| |
|`RefreshWorkbookLinksBlocked`|The operation failed because the user hasn't granted permission to refresh external workbook links.| |
|`RequestAborted`|The request was aborted during run time.| |
|`RequestPayloadSizeLimitExceeded`|The request payload size has exceeded the limit. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.| This error only occurs in Excel on the web.|
|`ResponsePayloadSizeLimitExceeded`|The response payload size has exceeded the limit. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.|  This error only occurs in Excel on the web.|
|`ServiceNotAvailable`|The service is unavailable.| |
|`Unauthenticated` |Required authentication information is either missing or invalid.| |
|`UnsupportedFeature`|The operation failed because the source worksheet contains one or more unsupported features.| |
|`UnsupportedOperation`|The operation being attempted is not supported.| |
|`UnsupportedSheet`|This sheet type does not support this operation, since it is a Macro or Chart sheet.| |

> [!NOTE]
> The preceding table lists error messages you may encounter while using the Excel JavaScript API. If you are working with the Common API instead of the application-specific Excel JavaScript API, see [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md) to learn about relevant error messages.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error object (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md)
