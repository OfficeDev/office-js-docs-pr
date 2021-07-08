---
title: Error handling with the Excel JavaScript API
description: 'Learn about Excel JavaScript API error handling logic to account for runtime errors.'
ms.date: 01/15/2021
localization_priority: Normal
---


# Error handling with the Excel JavaScript API

When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.

> [!NOTE]
> For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).

## Best practices

Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
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

|Error code | Error message |
|:----------|:--------------|
|`AccessDenied` |You cannot perform the requested operation.|
|`ActivityLimitReached`|Activity limit has been reached.|
|`ApiNotAvailable`|The requested API is not available.|
|`ApiNotFound`|The API you are trying to use could not be found. It may be available in a newer version of Excel. See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.|
|`BadPassword`|The password you supplied is incorrect.|
|`Conflict`|Request could not be processed because of a conflict.|
|`ContentLengthRequired`|A `Content-length` HTTP header is missing.|
|`GeneralException`|There was an internal error while processing the request.|
|`InactiveWorkbook`|The operation failed because multiple workbooks are open and the workbook being called by this API has lost focus.|
|`InsertDeleteConflict`|The insert or delete operation attempted resulted in a conflict.|
|`InvalidArgument` |The argument is invalid or missing or has an incorrect format.|
|`InvalidBinding`  |This object binding is no longer valid due to previous updates.|
|`InvalidOperation`|The operation attempted is invalid on the object.|
|`InvalidReference`|This reference is not valid for the current operation.|
|`InvalidRequest`  |Cannot process the request.|
|`InvalidSelection`|The current selection is invalid for this operation.|
|`ItemAlreadyExists`|The resource being created already exists.|
|`ItemNotFound` |The requested resource doesn't exist.|
|`NonBlankCellOffSheet`|Microsoft Excel can't insert new cells because it would push non-empty cells off the end of the worksheet. These non-empty cells might appear empty but have blank values, some formatting, or a formula. Delete enough rows or columns to make room for what you want to insert and then try again.|
|`NotImplemented`|The requested feature isn't implemented.|
|`RangeExceedsLimit`|The cell count in the range has exceeded the maximum supported number. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.|
|`RequestAborted`|The request was aborted during run time.|
|`RequestPayloadSizeLimitExceeded`|The request payload size has exceeded the limit. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information. <br><br>This error only occurs in Excel on the web.|
|`ResponsePayloadSizeLimitExceeded`|The response payload size has exceeded the limit. See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.  <br><br>This error only occurs in Excel on the web.|
|`ServiceNotAvailable`|The service is unavailable.|
|`Unauthenticated` |Required authentication information is either missing or invalid.|
|`UnsupportedOperation`|The operation being attempted is not supported.|
|`UnsupportedSheet`|This sheet type does not support this operation, since it is a Macro or Chart sheet.|

> [!NOTE]
> The preceding table lists error messages you may encounter while using the Excel JavaScript API. If you are working with the Common API instead of the application-specific Excel JavaScript API, see [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md) to learn about relevant error messages.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error object (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md)
