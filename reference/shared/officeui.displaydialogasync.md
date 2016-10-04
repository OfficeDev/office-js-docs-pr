# UI.displayDialogAsync method

Displays a dialog box in an Office host. 

## Requirements

|Host|Introduced in|Last changed in|
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

This method is available in the DialogAPI [requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md). To specify the DialogAPI requirement set, use the following in your manifest.

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

To detect this API at runtime, use the following code.

```js
 if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) 
 	{  
    	 // Use Office UI methods; 
 	} 
 else 
	 { 
	     // Alternate path 
	 } 
```



### Supported platforms
The DialogAPI requirement set is currently supported on the following platforms:

  - Office for Windows Desktop 2016 (build 16.0.6741.0000 or later)
  - Office for IPad (build 1.22 or later)
  - Office for Mac (build 15.20 or later) 

More platforms are coming soon. 

## Syntax

```js
office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##Examples

For a simple example that uses the **displayDialogAsync** method, see [Office Add-in Dialog API example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/) on GitHub.

For an example that shows an authentication scenario, see the [Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth) sample on GitHub.

 
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
|**width**|object|Optional. Defines the width of the dialog box as a percentage of the current display. The default value is 80%. The minimum resolution is 250 pixels.|
|**height**|object|Optional. Defines the height of the dialog box as a percentage of the current display. The default value is 80%. The minimum resolution is 150 pixels.|
|**displayInIFrame**|object|Optional. Determines whether the dialog box should be displayed within an IFrame in Office Online clients. This setting is ignored by desktop clients. The following are the possible values:<ul><li>False (default) - The dialog will be displayed as a new browser window (pop-up). Recommended for authentication pages that cannot be displayed in an IFrame. </li><li>True - The dialog will be displayed as a floating overlay with an IFrame. This is best for user experience and performance.</li>|


## Callback value
When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **displayDialogAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access the [Dialog](../../reference/shared/officeui.dialog.md) object.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined object or value, if you passed one as the _asyncContext_ parameter.|


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
