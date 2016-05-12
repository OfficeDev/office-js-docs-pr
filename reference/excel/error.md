# OfficeExtension.Error object (JavaScript API for Excel)

Represents errors that occur when you use the Excel JavaScript API.

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|code|string|Gets a value that indicates the type of error. The value can be "AccessDenied", "ActivityLimitReached", "BadPassword", "GeneralException", "InsertDeleteConflict", "InvalidArgument", "InvalidBinding", "InvalidOperation", "InvalidReference", "InvalidSelection", "ItemAlreadyExists", "ItemNotFound", "NotImplemented", or "UnsupportedOperation". |
|debugInfo|string|Gets a value that indicates what happened when the error occurred. This value is only intended for use during development / debugging.  |
|message |string| Gets a localized human readable string that corresponds to the error code.|
|name |string| Gets a value that is always "OfficeExtension.Error". |
|traceMessages |string[]| Gets an array of values that correspond to the instrumention messages set with context.trace(); |

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Returns the error code and message values in the following format: "{0}: {1}", code, message.|

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
