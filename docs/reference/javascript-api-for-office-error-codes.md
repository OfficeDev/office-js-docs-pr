---
title: Office Common API error codes
description: This article documents the error messages you might encounter while using the Office Common API.
ms.date: 09/01/2022
ms.localizationpriority: medium
---

# Office Common API error codes

This article documents the error messages you might encounter while using the Common API model. These error codes don't apply to application-specific APIs, such as the Excel JavaScript API or the Word JavaScript API.

See [API models](../develop/understanding-the-javascript-api-for-office.md#api-models) to learn more about the differences between the Common API and application-specific API models.

## Error codes

The following table lists the error codes, names, and messages displayed, and the conditions they indicate.

|Error.code|Error.name|Error.message|Condition|
|:-----|:-----|:-----|:-----|
|1000|Invalid Coercion Type|The specified coercion type is not supported|The coercion type is not supported in the Office application. (For example, OOXML and HTML coercion types are not supported in Excel.)|
|1001|Data Read Error|The current selection is not supported.|The user's current selection is not supported (that is, it is something different than the supported coercion types).|
|1002|Invalid Coercion Type|The specified coercion type is not compatible for this binding type.|The solution developer provided an incompatible combination of coercion type and binding type.|
|1003|Data Read Error|The specified rowCount or columnCount values are invalid.|The user supplies invalid column or row counts.|
|1004|Data Read Error|The current selection is not compatible for the specified coercion type.|The current selection is not supported for the specified coercion type by this application.|
|1005|Data Read Error|The specified startRow or startColumn values are invalid.|The user supplies invalid startRow or startCol values.|
|1006|Data Read Error|Coordinate parameters cannot be used with coercion type "Table" when the table contains merged cells.|The user tries to get partial data from a non-uniform table (that is, a table that has merged cells). |
|1007|Data Read Error|The size of the document is too large.|The user tries to get a document larger than the size currently supported.|
|1008|Data Read Error|The requested data set is too large.|The user requests to read data beyond the data limits defined by the Office application.|
|1009|Data Read Error|The specified file type is not supported.|The user sends an invalid file type.|
|2000|Data Write Error|The supplied data object type is not supported. |An unsupported data object is supplied.|
|2001|Data Write Error|Cannot write to the current selection.|The user's current selection is not supported for a write operation. (For example, when the user selects an image.)|
|2002|Data Write Error|The supplied data object is not compatible with the shape or dimensions of the current selection.|Multiple cells are selected (and the selection shape does not match the shape of the data). Multiple cells are selected (and the selection dimensions do not match the dimensions of the data).|
|2003|Data Write Error|The set operation failed because the supplied data object will overwrite data.|A single cell is selected and the supplied data object overwrites data in the worksheet.|
|2004|Data Write Error|The supplied data object does not match the size of the current selection.|The user supplies an object larger than the current selection size.|
|2005|Data Write Error|The specified startRow or startColumn values are invalid.|The user supplies invalid startRow or startCol values.|
|2006|Invalid Format Error|The format of the specified data object is not valid.|The solution developer supplies an invalid HTML or OOXML string, a malformed HTML string, or an invalid OOXML string.|
|2007|Invalid Data Object|The type of the specified data object is not compatible with the current selection.|The solution developer supplies a data object not compatible with the specified coercion type.|
|2008|Data Write Error|TBD|TBD|
|2009|Data Write Error|The specified data object is too large.|The user tries to set data beyond the data limits defined by the Office application.|
|2010|Data Write Error|Coordinate parameters cannot be used with coercion type Table when the table contains merged cells.|The user tries to set partial data from a non- uniform table (that is, a table that has merged cells).|
|3000|Binding Creation Error|Cannot bind to the current selection.|The user's selection is not supported for binding. (For example, the user is selecting an image or other non-supported object.)|
|3001|Binding Creation Error|TBD|TBD|
|3002|Invalid Binding Error|The specified binding does not exist.|The developer tries to bind to a non-existing or removed binding.|
|3003|Binding Creation Error|Noncontiguous selections are not supported.|The user is making multiple selections.|
|3004|Binding Creation Error|A binding cannot be created with the current selection and the specified binding type.|There are several conditions under which this might happen. Please see the "Binding creation error conditions" section later in this article.|
|3005|Invalid Binding Operation|Operation is not supported on this binding type.|The developer sends an add row or add column operation on a binding type that is not  of coercion type `table`.|
|3006|Binding Creation Error|The named item does not exist.|The named item cannot be found. No content control or table with that name exists.|
|3007|Binding Creation Error|Multiple objects with the same name were found.|Collision error: more than one content control with the same name exists, and fail on collision is set to `true`.|
|3008|Binding Creation Error|The specified binding type is not compatible with the supplied named item.|Named item can't be bound to type. For example, a content control contains text, but the developer tried to bind by using coercion type `table`.|
|3009|Invalid Binding Operation|The binding type is not supported.|Used for backward compatibility.|
|3010|Unsupported Binding Operation|The selected content needs to be in table format. Format the data as a table and try again.|The developer is trying to use the `addRowsAsync` or `deleteAllDataValuesAsync` method of the `TableBinding` object on data of coercion type `matrix`.|
|4000|Read Settings Error|The specified setting name does not exist.|A nonexistent setting name is supplied.|
|4001|Save Settings Error|The settings could not be saved.|Settings could not be saved.|
|4002|Settings Stale Error|Settings could not be saved because they are stale.|Settings are stale and developer indicated not to override settings.|
|5000|Settings Stale Error|The operation is not supported.|The operation is not supported in the current Office application. For example, `document.getSelectionAsync` is called from Outlook.|
|5001|Internal Error|An internal error has occurred.|Refers to an internal error condition, which can occur for any of the following reasons.<br/><table><tr><td>An add-in being used by another user sharing the workbook created a binding at approximately the same time, and your add-in needs to retry binding.</tr></td><tr><td>An unknown error occurred.</tr></td><tr><td>The operation failed.</tr></td><tr><td>Access was denied because the user is not a member of an authorized role.</tr></td><tr><td>Access was denied because secure, encrypted communication is required.</tr></td><tr><td>Data is stale and the user needs to confirm enabling the queries to refresh it.</tr></td><tr><td>The site collection CPU quota has been exceeded.</tr></td><tr><td>The site collection memory quota has been exceeded.</tr></td><tr><td>The session memory quota has been exceeded.</tr></td><tr><td>The workbook is in an invalid state and the operation can't be performed.</tr></td><tr><td>The session has timed out due to inactivity and the user needs to reload the workbook.</tr></td><tr><td>The maximum number of allowed sessions per user has been exceeded.</tr></td><tr><td>The operation was canceled by the user.</tr></td><tr><td>The operation can't be completed because it is taking too long.</tr></td><tr><td>The request can't be completed and needs to be retried.</tr></td><tr><td>The trial period of the product has expired.</tr></td><tr><td>The session has timed out due to inactivity.</tr></td><tr><td>The user doesn't have permission to perform the operation on the specified range.</tr></td><tr><td>The user's regional settings don't match the current collaboration session.</tr></td><tr><td>The user is no longer connected and must refresh or re-open the workbook.</tr></td><tr><td>The requested range doesn't exist in the sheet.</tr></td><tr><td>The user doesn't have permission to edit the workbook.</tr></td><tr><td>The workbook can't be edited because it is locked.</tr></td><tr><td>The session can't auto save the workbook.</tr></td><tr><td>The session can't refresh its lock on the workbook file.</tr></td><tr><td>The request can't be processed and needs to be retried.</tr></td><tr><td>The user's sign-in information couldn't be verified and needs to be re-entered.</tr></td><tr><td>The user has been denied access.</tr></td><tr><td>The shared workbook needs to be updated.</tr></td></table>|
|5002|Permission Denied|The requested operation is not allowed on the current document mode.|The solution developer submits a set operation, but the document is in a mode that does not allow modifications, such as 'Restrict Editing'.|
|5003|Event Registration Error|The specified event type is not supported by the current object.|The solution developer tries to register or unregister a handler to an event that does not exist.|
|5004|Invalid API call|Invalid API call in the current context.|An invalid call is made for the context, for example, trying to use a `CustomXMLPart` object in Excel.|
|5005|Data Stale|Operation failed because the data is stale on the server.|The data on the server needs to be refreshed.|
|5006|Session Timeout|The document session timed out. Reload the document. |The session has timed out.|
|5007|Invalid API call|The enumeration is not supported in the current context.|The enumeration is not supported in the current context.|
|5009|Permission Denied|Access Denied|The add-in does not have permission to call the specific API.|
|5012|Invalid Or Timed Out Session|Your Office browser session has expired or is invalid. To continue, refresh the page.|The session between the Office client and server has expired or the date, time, or time zone is incorrect on your computer.|
|6000|Invalid node|The specified node was not found.|The `CustomXmlPart` node was not found.|
|6100|Custom XML error|Custom XML error|Invalid API call.|
|7000|Invalid Id|The specified Id does not exist.|Invalid ID.|
|7001|Invalid navigation|The object is located in a place where navigation is not supported.|The user can find the object, but cannot navigate to it. (For example, in Word, the binding is to the header, footer, or a comment.)|
|7002|Invalid navigation|The object is locked or protected.|The user is trying to navigate to a locked or protected range.|
|7004|Invalid navigation|The operation failed because the Index is out of range.|The user is trying to navigate to an index that is out of range.|
|8000|Missing Parameter|We couldn't format the table cell because some parameter values are missing. Double-check the parameters and try again.|The cellFormat method is missing some parameters. For example, there are missing cells, format, or tableOptions parameters.|
|8010|Invalid value|One or more of the cells parameters have values that aren't allowed. Double-check the values and try again.|The common cells reference enumeration is not defined. For example, All, Data, Headers.|
|8011|Invalid value|One or more of the tableOptions parameters have values that aren't allowed. Double-check the values and try again.|One of the values in tableOptions is invalid.|
|8012|Invalid value|One or more of the format parameters have values that aren't allowed. Double-check the values and try again.|One of the values in the format is invalid.|
|8020|Out of range|The row index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of rows.|The row index is more than the biggest row index of the table or less than 0.|
|8021|Out of range|The column index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of columns.|The column index is more than the biggest column index of the table or less than 0.|
|8022|Out of range|The value is out of the allowed range.|Some of the values in the format are out of the supported ranges.|
|9016|Permission denied|Permission denied|Access is denied.|
|9020|Generic Response Error|An internal error has occurred.|Refers to an internal error condition, which can occur for any number of reasons.|
|9021|Save Error|Connection error occurred while trying to save the item on the server.|The item couldn't be saved. This could be due to a server connection error if using Online Mode in Outlook desktop, or due to an attempt to re-save a draft item that was deleted from the Exchange server.|
|9022|Message In Different Store Error|The EWS ID cannot be retrieved because the message is saved in another store.|The EWS ID for the current message couldn't be retrieved as the message may have been moved or the sending mailbox may have changed.|
|9041|Network error|The user is no longer connected to the network. Please check your network connection and try again.|The user no longer has network or internet access.|
|9043|Attachment Type Not Supported|The attachment type is not supported.|The API doesn't support the attachment type. For example, `item.getAttachmentContentAsync` throws this error if the attachment is an embedded image in Rich Text Format, or if it's an item type other than an email or calendar item (such as a contact or task item).|
|9057|Size Limit Exceeded|A maximum of 32KB is available for the settings of each add-in.|When updating roaming settings via Office.context.roamingSettings.set, the size cannot exceed 32KB. See [Office.RoamingSettings interface](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-set-member(1)).
|12002|*Not applicable*|*Not applicable*|One of the following:<br> - No page exists at the URL that was passed to `displayDialogAsync`.<br> - The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax. Thrown within the dialog and triggers a `DialogEventReceived` event in the host page.|
|12003|*Not applicable*|*Not applicable*|The dialog box was directed to a URL with the HTTP protocol. HTTPS is required. Thrown within the dialog and triggers a `DialogEventReceived` event in the host page.|
|12004|*Not applicable*|*Not applicable*|The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number). Thrown by call of `displayDialogAsync`.|
|12005|*Not applicable*|*Not applicable*|The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. Thrown by call of `displayDialogAsync`. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)|
|12006|*Not applicable*|*Not applicable*|The dialog box was closed, usually because the user chooses the **X** button. Thrown within the dialog and triggers a `DialogEventReceived` event in the host page.|
|12007|*Not applicable*|*Not applicable*|A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time. Thrown by call of `displayDialogAsync`.|
|12009|*Not applicable*|*Not applicable*|The user chose to ignore the dialog box. This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog. Thrown by call of `displayDialogAsync`.|
|12011|*Not applicable*|*Not applicable*|The user's browser is configured in a way that blocks popups. This error can occur in Office on the web if the browser is Safari and it's configured to block popups or the browser is Edge Legacy (an older, unsupported version of the current Chromium-based Edge) and the add-in domain is in a different security zone from the domain the dialog is trying to open. Thrown by call of `displayDialogAsync`.|
|13nnn|*Not applicable*|*Not applicable*|See [Causes and handling of errors from getAccessToken](../develop/troubleshoot-sso-in-office-add-ins.md#causes-and-handling-of-errors-from-getaccesstoken).|

## Binding creation error conditions

When a binding is created in the API, indicate the binding type that you want to use. The following tables lists the binding types and the resulting binding behaviors that are expected.

### Behavior in Excel

The following table summarizes binding behavior in Excel.

|Specified Binding Type|Actual Selection|Behavior|
|:-----|:-----|:-----|
|Matrix|Range of cells (including within a table, and single cell)|A binding of type `matrix` is created on the selected cells. No modification in the document is expected.|
|Matrix|Text selected in the cell|A binding of type `matrix` is created on the whole cell. No modification in the document is expected.|
|Matrix|Multiple selection/invalid selection (For example, user selects a picture, object, or Word Art.)|The binding cannot be created.|
|Table|Range of cells (includes single cell)|The binding cannot be created.|
|Table|Range of cell within a table (includes single cell within a table, or the whole table, or text within a cell in a table)|A binding is created in the whole table.|
|Table|Half selection in a table and half selection outside the table|The binding cannot be created.|
|Table|Text selected in the cell (not in the table.)|The binding cannot be created.|
|Table|Multiple selection/invalid selection (For example, user selects a picture, object, Word Art, etc.)|The binding cannot be created.|
|Text|Range of cells|The binding cannot be created.|
|Text|Range of cells within a table|The binding cannot be created.|
|Text|Single cell|A binding of type `text` is created.|
|Text|Single cell within a table|A binding of type `text` is created.|
|Text|Text selected in the cell|A binding of type `text` in the whole cell is created.|

### Behavior in Word

The following table summarizes binding behavior in Word.

|Specified Binding Type|Actual Selection|Behavior|
|:-----|:-----|:-----|
|Matrix|Text|The binding cannot be created.|
|Matrix|Whole table|A binding of type `matrix` is created.Document is changed and a content control must wrap the table. |
|Matrix|Range within a table|The binding cannot be created.|
|Matrix|Invalid selection (for example, multiple, invalid objects, etc.)|The binding cannot be created.|
|Table|Text|The binding cannot be created.|
|Table|Whole table|A binding of type `text` is created.|
|Table|Range within a table|The binding cannot be created.|
|Table|Invalid selection (for example, multiple, invalid objects, etc.)|The binding cannot be created.|
|Text|Whole table|A binding of type `text` is created.|
|Text|Range within a table|The binding cannot be created.|
|Text|Multiple selection|The last selection will be wrapped with a content control and a binding to that control. A content control of type `text` is created.|
|Text|Invalid selection (for example, multiple, invalid objects, etc.)|The binding cannot be created.|

## See also

- [Office Add-ins development lifecycle](../overview/office-add-ins.md)
- [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md)
- [Error handling with the application-specific JavaScript APIs](../testing/application-specific-api-error-handling.md)
- [Troubleshoot error messages for single sign-on (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md)
