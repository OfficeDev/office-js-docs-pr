# OfficeExtension.Error Object (JavaScript API for OneNote)

Represents errors that occur when you use the OneNote API.

_Applies to: OneNote Online_
_Note: This API is in preview_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|code|string|Gets a value that indicates the type of error. The value can be "AccessDenied", "GeneralException", or "ActivityLimitReached". |
|debugInfo|string|Gets a value that indicates what happened when the error occurred. This value is only intended for use during development / debugging.  |
|message |string| Gets a localized human readable string that corresponds to the error code.|
|name |string| Gets a value that is always "OfficeExtension.Error". |
|traceMessages |string[]| Gets an array of values that correspond to the instrumention messages set with context.trace(); |

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[toString()](#toString)|string|Returns the error code and message values in the following format: "{0}: {1}", code, message.|

## Method details

### toString()
Returns the error code and message values in the following format: "{0}: {1}", code, message.

#### Syntax
```js
error.toString()
```

#### Parameters
None.

#### Returns
string
