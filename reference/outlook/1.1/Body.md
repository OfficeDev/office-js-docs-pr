

# Body

The `body` object provides methods for adding and updating the content of the message or appointment. It is returned in the `body` property of the selected item.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Methods

####  getTypeAsync([options], [callback])

Gets a value that indicates whether the content is in HTML or text format.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

The content type is returned as one of the [CoercionType](Office.md#coerciontype-string) values in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|
####  prependAsync(data, [options], [callback])

Adds the specified content to the beginning of the item body.

The `prependAsync` method inserts the specified string at the beginning of the item body. Calling the `prependAsync` method is the same as calling the [`setSelectedDataAsync`](Body.md#setselecteddataasyncdata-options-callback) method with the insertion point at the beginning of the body content.

When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (`<a>`) to `LPNoLP`. For example:

```
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`data`| String||The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`options.coercionType`| [Office.CoercionType](Office.md)| &lt;optional&gt;|The desired format for the body. The string in the `data` parameter will be converted to this format.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>Any errors encountered will be provided in the `asyncResult.error` property.|

##### Errors

| Error code | Description |
|------------|-------------|
| `DataExceedsMaximumSize` | The `data` parameter is longer than 1,000,000 characters. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|
####  setSelectedDataAsync(data, [options], [callback])

Replaces the selection in the body with the specified text.

The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the UI, the string will be inserted at the top of the body content.

When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (`<a>`) to `LPNoLP`. For example:

```
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`data`| String||The string to be inserted in the body. The string is limited to 1,000,000 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`options.coercionType`| [Office.CoercionType](Office.md)| &lt;optional&gt;|The desired format for the body. The string in the `data` parameter will be converted to this format.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>Any errors encountered will be provided in the `asyncResult.error` property.|

##### Errors

| Error code | Description |
|------------|-------------|
| `DataExceedsMaximumSize` | The `data` parameter is longer than 1,000,000 characters. |
| `InvalidFormatError` | The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|
