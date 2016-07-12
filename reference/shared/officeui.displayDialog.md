# UI.displayDialog method

Displays a web dialog inside Office hosts. 

## Requirements

|Host|Introduced in|Last changed in|
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

To require the `DialogAPI` [requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md) 1.1 or later, your manifest should specify

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

Runtime detection of the this API can be done with the following code:

```js
 if (Office.context.requirements.isSetSupported('DialogAPI', '1.1')) 
 	{  
    	 // Use Office UI methods; 
 	} 
 else 
	 { 
	     // Alternate path 
	 } 
```



###Supported platforms
The Dialog API is currently supported on the following platforms:

  - Office for Windows Desktop 2016 (build 16.0.6741.0000 or above)
  - Office for IPad (build 1.22 or above)
  - Office for Mac (build 15.20 or above) 
  - More platforms coming soon. 

## Syntax
```js
office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##Examples

The following examples illustrate the use of the dialog API


- **Simple use**: [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/)
- **Authentication**: [Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)

 
## Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startAddress|string|Accepts the initial HTTPS(TLS) Url that opens in the dialog. <ul><li>The initial page must be on the same domain as the parent page. After the initial page is loaded in the dialog you can then navigate to other domains as long as they are declared on the AppDomains list on the manifest.</li> <li>Any page calling [office.context.ui.messageParent](officeui.messageParent.md) must also be on the same domain as the parent page.</li></ul>|
|options|object|Optional. Accepts an options object to define dialog behaviors.|
|callback|object|Accepts a callback method to handle the dialog creation attempt.|
	
### Dialog Options
Dialogs support a number of configuration options.


| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|width|object|Optional. Defines the width of the dialog as a percentage of the current display. Defaults to 80%. 250px minimum.|
|height|object|Optional. Defines the height of the dialog as a percentage of the current display. Defaults to 80%. 150px minimum.|
|displayAsIFrame|object|Optional. Determines whether the dialog should be displayed within an IFrame in Online Clients (e.g. Word Online). This setting is ignored in non Online Clients. <ul><li>False (default): The dialog will be displayed as a new browser window (pop-up); recommended for authentication pages that cannot be IFramed. </li><li>True: The dialog will be displayed as a floating overlay with an IFrame; this is best for user experience and performance. Ensure that any pages displayed on the dialog can be IFramed. </li>|


## Callback value
When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **displayDialogAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access the **[Dialog](../../reference/shared/officeui.dialog.md)** object.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|


	
## Design considerations
###Remarks
- An Office add-in may have only 1 dialog open at any time 
- Every dialog can be moved and resized by the user
- Every dialog opens centered on the screen 
- Dialogs appear on top of the app and one another in order of being created

###Use a dialog to
- Display authentication pages to collect user credentials 
- Display an error/progress/input screen from a ShowTaspane or ExecuteAction command
- Temporarily increase the real state the user needs to achieve a task

###Do not use a dialog to
- Interact back and forth with a document. Use a taskpane instead. 

###Useful patterns
See [Client Dialog](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md) in Office Add-in UX Design Patterns
