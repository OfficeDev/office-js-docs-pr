#UI.Dialog Object
This is the object that gets returned after calling [office.context.ui.displayDialog](officeui.displayDialog.md)

## Members
| Member	   | Type	|Description|
|:---------------|:--------|:----------|
|close|function|Allows the add-in to close its dialog.|
|DialogMessageReceived|event|Optional. Triggered when the dialog sends a message.|
|DialogEventReceived|event|Optional. Triggered when the dialog has been closed or otherwise unloaded.|


### close()
Called from a parent page it closes the corresponding dialog.     
```js    
[dialogObject].close();    
``` 

#### Parameters    
None. 

#### Returns    
void  


#### Examples
See full examples on the [displayDialog](officeui.displayDialog.md) reference page