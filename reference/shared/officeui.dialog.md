#UI.Dialog object
The object that is returned when the [displayDialogAsync](officeui.displaydialogasync.md) method is called.

## Members
| Member	   | Type	|Description|
|:---------------|:--------|:----------|
|close|Function|Allows the add-in to close its dialog box.|
|addEventHandler|Function|Registers an event handler. The two supported events are: <ul><li>DialogMessageReceived. Triggered when the dialog box sends a message to its parent.</li><li>DialogEventReceived. Triggered when the dialog box has been closed or otherwise unloaded.</li></ul> |


### close()
Called from a parent page to close the corresponding dialog box.     
```js    
[dialogObject].close();    
``` 

#### Parameters    
None. 

#### Returns    
void  


#### Examples
For examples, see the [DisplayDialogAsync method](officeui.displaydialogasync.md) topic.
