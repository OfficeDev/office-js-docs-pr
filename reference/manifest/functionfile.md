# FunctionFile element

Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [FormFactor](./formfactor). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons. For more information, see [Button](./button.md).

The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`. The functions should use the [item.notificationMessages](../../../reference/outlook/Office.context.mailbox.item.md) API to indicate progress, success, or failure to the user. It should also call [event.completed](../../../reference/shared/event.completed.md) when it has finished execution. The name of the functions are used in the **FunctionName** element for UI-less buttons.

The following is an example of an HTML file defining a **trackMessage** function.

```js
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```