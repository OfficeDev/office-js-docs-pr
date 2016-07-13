#UI.Dialog object
The object that is returned when the [displayDialog](officeui.displayDialog.md) method is called.

## Members
| Member	   | Type	|Description|
|:---------------|:--------|:----------|
|close|Function|Allows the add-in to close its dialog box.|
|DialogMessageReceived|Event|Optional. Triggered when the dialog box sends a message.|
|DialogEventReceived|Event|Optional. Triggered when the dialog box has been closed or otherwise unloaded.|


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
For examples, see the [DisplayDialog method](officeui.displayDialog.md) topic.
