# OfficeExtension.Error Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents errors that occur when you use the OneNote JavaScript API.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-error).

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|code|string|Gets a value that indicates the type of error. The value can be **InvalidArgument**, **GeneralException**, **ItemNotFound** or **UnsupportedOperationForObjectType**. |
|debugInfo|string|Gets a value that indicates what happened when the error occurred. This value is only intended for use during development and debugging.  |
|message |string| Gets a localized human readable string that corresponds to the error code.|
|name |string| Gets a value that is always **OfficeExtension.Error**. |
|traceMessages |string[]| Gets an array of values that correspond to the instrumention messages set with `context.trace();`. |


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Returns the error code and message values in the following format: `"{0}: {1}", code, message`.|



## Method details

### toString()

Returns the error code and message values in the following format: `"{0}: {1}", code, message`.

#### Syntax

```js
error.toString()
```

#### Parameters

None

#### Returns

String
