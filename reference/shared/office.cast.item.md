
# Office.cast.item property
Provides IntelliSense specific to compose or read mode messages and appointments.

|||
|:-----|:-----|
|**Hosts:**|Outlook|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Mailbox|
|**Last changed in**|1.0|



|||
|:-----|:-----|
|**Applicable Outlook modes**|Design time in Visual Studio only|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## Return value

A set of methods that enable you to select the appropriate IntelliSense for your Outlook add-in.


## Remarks

This property and its methods support IntelliSense for developing Outlook add-ins on Visual Studio only. They do not have any effect on other development tools.

The  **Office.cast.item** methods are used at design time in Visual Studio to provide specific IntelliSense for the **Office.context.mailbox.item** property. When you use the **toAppointmentCompose** method, for example, IntelliSense will show only the **Appointment** methods and properties that apply in compose mode.

At run time, the  **Office.cast.item** methods have no effect on your Outlook add-in.


## Example

The following example uses the  **toMessageCompose** method to cast the **Office.context.mailbox.item** property so that it will only show IntelliSense for the **Message** object in compose mode. After the cast, the `message` variable will only display IntelliSense for methods and properties that can be used in compose mode.


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).

||Office for Windows desktop|Office Online (in browser)|Outlook for Mac|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|
