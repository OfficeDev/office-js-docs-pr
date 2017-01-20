# UI.displayDialogAsync method

Displays a dialog box in an Office host. 

## Requirements

|Host|Introduced in|Last changed in|
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

This method is available in the DialogAPI [requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md) for Word, Excel, or PowerPoint add-ins, and in the Mailbox requirement set 1.4 for Outlook. To specify the DialogAPI requirement set, use the following in your manifest.

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.1"> 
    <Set Name="DialogAPI"/> 
  </Sets> 
</Requirements> 
```

To specify the Mailbox 1.4 requirement set, use the following in your manifest.

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.4"> 
    <Set Name="Mailbox"/> 
  </Sets> 
</Requirements> 
```

To detect this API at runtime in a Word, Excel, or PowerPoint add-in, use the following code.

```js
if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

To detect this API at runtime in an Outlook add-in, use the following code.

```js
if (Office.context.requirements.isSetSupported('Mailbox', 1.4)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

Alternatively, you can check if the `displayDialogAsync` method is undefined before using it.

```js
if (Office.context.ui.displayDialogAsync !== undefined) {
  // Use Office UI methods
}
```

### Supported platforms
For information about supported platforms, see [Dialog API requirement sets](../requirement-sets/dialog-api-requirement-sets.md).

## Syntax

```js
Office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##Examples

For a simple example that uses the **displayDialogAsync** method, see [Office Add-in Dialog API example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/) on GitHub.

For an examples that show authentication scenarios, see:

- [PowerPoint Add-in in Microsoft Graph ASP.Net Insert Chart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Office Add-in Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Excel Add-in ASP.NET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Office Add-in Server Authentication Sample for ASP.net MVC](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)
- [Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)


 
## Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startAddress|string|Accepts the initial HTTPS(TLS) URL that opens in the dialog box. <ul><li>The initial page must be on the same domain as the parent page. After the initial page loads, you can go to other domains.</li><li>Any page calling [office.context.ui.messageParent](officeui.messageparent.md) must also be on the same domain as the parent page.</li></ul>|
|options|object|Optional. Accepts an options object to define dialog behaviors.|
|callback|object|Accepts a callback method to handle the dialog creation attempt.|
	
### Configuration options
The following configuration options are available for a dialog box.


| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|**width**|int|Optional. Defines the width of the dialog box as a percentage of the current display. The default value is 80%. The minimum resolution is 250 pixels.|
|**height**|int|Optional. Defines the height of the dialog box as a percentage of the current display. The default value is 80%. The minimum resolution is 150 pixels.|
|**displayInIframe**|bool|Optional. Determines whether the dialog box should be displayed within an IFrame. **This setting is only applicable in Office Online clients**, this setting is ignored by desktop clients. The following are the possible values:<ul><li>false (default) - The dialog will be displayed as a new browser window (pop-up). Recommended for authentication pages that cannot be displayed in an IFrame. </li><li>true - The dialog will be displayed as a floating overlay with an IFrame. This is best for user experience and performance.</li>|


## Callback value
When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **displayDialogAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access the [Dialog](../../reference/shared/officeui.dialog.md) object.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined object or value, if you passed one as the _asyncContext_ parameter.|

### Errors from displayDialogAsync

In addition to general platform and system errors, the following errors are specific to calling **displayDialogAsync**.

|**Code number**|**Meaning**|
|:-----|:-----|
|12004|The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be either the same domain as the host page (including protocol and port number), or it must be registered in the `<AppDomains>` section of the add-in manifest.|
|12005|The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)|
|12007|A dialog box is already opened from the task pane. A task pane add-in can only have one dialog box open at a time.|



## Design considerations
The following design considerations apply to dialog boxes:

- An Office Add-in can have only one dialog box open at any time.
- Every dialog box can be moved and resized by the user.
- Every dialog box is centered on the screen when opened.
- Dialog boxes appear on top of the host application and in the order in which they were created.

Use a dialog box to:

- Display authentication pages to collect user credentials.
- Display an error/progress/input screen from a ShowTaspane or ExecuteAction command.
- Temporarily increase the surface area that a user has available to complete a task.

Do not use a dialog box to interact with a document. Use a task pane instead. 

For a design pattern that you can use to create a dialog box, see [Client Dialog](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md) in the Office Add-in UX Design Patterns repository on GitHub.
