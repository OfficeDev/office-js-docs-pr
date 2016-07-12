# UI.messageParent method

### messageParent(messageObject: object)
Delivers a message from the dialog to its parent/opener page. The page calling this API must be on the same domain as the parent. 

#### Syntax
```js
Office.context.ui.messageParent("Message from Dialog");
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|messageObject|string or boolean|Accepts a message from the dialog to deliver to the add-in.|

#### Returns
void

#### Examples
See full examples on the [displayDialog](officeui.displayDialog.md) reference page
