# UI.messageParent method

Delivers a message from the dialog box to its parent/opener page. The page calling this API must be on the same domain as the parent. 

## Syntax

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|messageObject|String or Boolean|Accepts a message from the dialog box to deliver to the add-in.|

## Returns
void

## Examples
For examples, see the [DisplayDialogAsync method](officeui.displaydialogasync.md) topic.

