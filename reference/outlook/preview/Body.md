

# Body

The `body` object provides methods for adding and updating the content of the message or appointment. It is returned in the `body` property of the selected item.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Methods

####  getAsync(coercionType, [options], [callback])

Returns the current body in a specified format.

This method returns the entire current body in the format specified by `coercionType`.

When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent. The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method previously. The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||The format for the returned body.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. The body is provided in the requested format in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Examples

This example gets the body of the message in plain text.

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Do something with the result
  });
```

The following is an example of the `result` parameter passed to the callback function.

```js
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  getTypeAsync([options], [callback])

Gets a value that indicates whether the content is in HTML or text format.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. The content type is returned as one of the [CoercionType](Office.md#coerciontype-string) values in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|
####  prependAsync(data, [options], [callback])

Adds the specified content to the beginning of the item body.

The `prependAsync` method inserts the specified string at the beginning of the item body. Calling the `prependAsync` method is the same as calling the [`setSelectedDataAsync`](#setselecteddataasyncdata-options-callback) method with the insertion point at the beginning of the body content.

When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (`<a>`) to `LPNoLP`. For example:

```js
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
####  setAsync(data, [options], [callback])

Replaces the entire body with the specified text.

The `setAsync` method replaces the existing body of the item with the specified string or, if text is selected in the editor, it replaces the selected text.

When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent. The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method previously. The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.

When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (`<a>`) to `LPNoLP`. For example:

```js
Office.context.mailbox.item.body.setAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`data`| String||The string that will replace the existing body. The string is limited to 1,000,000 characters.|
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
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Examples

The following example replaces the body with HTML content.

```js
Office.context.mailbox.item.body.setAsync(
  "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
  { coercionType:"html", asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Process the result
  });
```

The following is an example of the `result` parameter passed to the callback function.

```js
{
  "value":null,
  "status":"succeeded",
  "asyncContext":"This is passed to the callback"
}
```

####  setSelectedDataAsync(data, [options], [callback])

Replaces the selection in the body with the specified text.

The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the UI, the string will be inserted at the top of the body content.

When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (`<a>`) to `LPNoLP`. For example:

```js
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
