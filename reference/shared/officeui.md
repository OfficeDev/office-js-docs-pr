# Office UI Namespace (JavaScript API for Office)

The Office UI Namespace, Office.context.ui, provides objects and methods used to create UI components for add-ins.

##### Requirements

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

Runtime detection of the `DialogAPI` capability can be done with the following code:

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

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[displayDialogAsync()](#displayDialogAsync(startAddress: Url, options: Object, callback: function))|void|Displays a dialog to display or collect information from the user or to facilitate Web navigation.|
|[messageParent()](#messageParent(messageObject: object))|void|Sends a message from a dialog to the parent add-in.|

## Method Details

### displayDialogAsync(startAddress: Url, options: Object, callback: function)
Display up to one Web dialog that opens at startAddress.

#### Syntax
```js
displayDialogAsync(startAddress, options, callback);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startAddress|string|Accepts the initial HTTPS Url that opens in the dialog.|
|options|object|Optional. Accepts an options object to define dialog behaviors.|
|callback|object|Accepts a callback method to handle the dialog creation attempt.|

#### Returns
void

#### Examples

```js
var dialog;

function dialogCallback(asyncResult){ 
	dialog = asyncResult.value; 
	dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, messageHandler); 
	dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived,  eventHandler); 
} 

function messageHandler(arg){ 
	actOnMessage(arg.message); 
} 

function eventHandler(arg){ 
	actOnEvent(arg.message);
	dialog.close(); 
} 

function openDialog() {
	Office.context.ui.displayDialogAsync("https://addin.app.net/dialog.html",  
		{height:80, width:50}, dialogCallback); 
}
```

#### Comments
1.	Dialogs will not create modal windows
2.	The initial url of the dialog and any page using the messageParent API must be on the same domain as the parent. 
3.	An Office add-in may have 1 dialog open at any time 
3.	Every dialog may be moved and resized
4.	Every dialog opens centered on the screen 
5.	Dialogs appear on top of the app and one another in order of being created
6.	Dialogs can only navigate to secured (TLS) sites 
7.	Dialogs must initially open to a site on the add-in manifest's App Domains list
8.	Dialogs cannot send messages from pages outside the add-in manifest's App Domains list

### callback()
The callback for displayDialogAsync, in the success case, includes a dialog object. This dialog object has additional behaviors. 

#### Dialog Object
| Member	   | Type	|Description|
|:---------------|:--------|:----------|
|close|function|Allows the add-in to close its dialog.|
|DialogMessageReceived|event|Optional. Triggered when the dialog sends a message.|
|DialogEventReceived|event|Optional. Triggered when the dialog has been closed or otherwise unloaded.|

#### Returns
void

### close()
When called from an active add-in, immediately closes its dialog. 

#### Syntax    
```js    
close();    
``` 

#### Parameters    
None. 

#### Returns    
void  

#### Examples    
    
```js    
//using _dlg object provided by the displayDialogAsync callback method    
closeButton.addEventListener("click", _dlg.close);    
```  

### messageParent(messageObject: object)
Delivers a message from the dialog to its parent add-in.

#### Syntax
```js
messageParent("Message from Dialog");
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|messageObject|object|Accepts a message from the dialog to deliver to the add-in.|

#### Returns
void

#### Examples

```js
messageParent("Message from Dialog");
```

## Objects

### Dialog Options
Dialogs support a number of configuration options.

#### Properties
| Properties	   | Type	|Description|
|:---------------|:--------|:----------|
|width|object|Optional. Defines the width of the dialog as a percentage of the current display. Defaults to 80%. 250px minimum.|
|height|object|Optional. Defines the height of the dialog as a percentage of the current display. Defaults to 80%. 150px minimum.|
|xFrameDenySafe|object|Optional. Determines whether the dialog is safe to display within a Web frame.|
|enforceAppDomains|object|Optional. Restricts the dialog's navigation to the add-in's trusted sites.|

#### Comments
1.	The default dialog dimensions are 80% display width x 80% display height (based on the current device dimensions) 
2.	Dialogs have a minimum size to avoid discoverability problems 
3.	The dialog display may be in portrait or landscape orientation, and the width and height will adjust accordingly 

## Supported platforms
The Dialog API is currently supported on the following platforms:

  - Office for Windows Desktop 2016 (build 16.0.6741.0000 or above)

More platforms coming soon. 
