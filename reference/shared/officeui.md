# Office UI Namespace (JavaScript API for Office)

_Applies to: Office Online, Office 2013, Office 2016_

The Office UI Namespace provides objects and methods used to create UI components for add-ins.


##### Requirements

Usage of Office UI APIs require `DialogAPI` requirement set 1.1 or later. To set this for app activation, use:

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

Runtime detection of `DialogAPI` capability can be done with the following code:

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
|[close()](#close)|void|Closes the dialog.|
|[displayDialogAsync()](#displaydialogasync)|void|Displays a dialog to display or collect information from the user or to facilitate Web navigation.|
|[messageParent()](#messageparent)|void|Sends a message from a dialog to the parent add-in.|

## Method Details

### close()
When called from an active add-in dialog, immediately closes the dialog.

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
closeButton.addEventListener("click", office.ui.close);
```

### displayDialogAsync(startAddress: string[, options: object], callback: function)
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
function dialogCallback(asyncResult){ 
	var dialog = asyncResult.value; 
	dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, messageHandler); 
	dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived,  eventHandler); 
} 

function messageHandler(arg){ 
	actOnMessage(arg.message); 
} 

function eventHandler(arg){ 
	actOnEvent(arg.message); 
} 

function openDialog() {
	Office.context.ui.displayDialogAsync("https://addin.app.net/dialog.html",  
		{height:80, width:50}, dialogCallback); 
}
```

#### Comments
1.	Dialogs will not create modal windows
2.	An Office add-in may have 1 dialog open at any time 
3.	Every dialog may be resized
4.	Every dialog opens centered on the screen 
5.	Dialogs always open on top
6.	Dialogs can only navigate to secured (TLS) sites 
7.	Dialogs must initially open a site on the Trusted Domains list
8.	Dialogs cannot send messages from pages outside the Trusted Domains list

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
|enforceAppDomains|object|Optional|object|Optional. Restricts the dialog's navigation to the add-in's trusted sites.|

#### Comments
1.	The default dialog dimensions are 80% display width x 80% display height (based on the current device dimensions) 
2.	Dialogs have a minimum size to avoid discoverability problems 
3.	The dialog display may be in portrait or landscape orientation, and the width and height will adjust accordingly