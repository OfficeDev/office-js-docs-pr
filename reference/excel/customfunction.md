# CustomFunction Object (JavaScript API for Excel)

Defines a custom function object in Excel.

## Properties

| Property	   | Type	|Description| 
|:---------------|:--------|:----------|:----|
|description|string|A description of the custom function. Read-only.|
|id|string|The ID of the function. Read-only.|
|name|string|The name of the function. Read-only.|
|resultType|string|Result type returned by the function. Read-only.|
|streaming|bool|Represents whether the function is running in streamed mode or not. Read-only.|

## Relationships
None

## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes this function from Excel.|

## Method Details

### delete()
Deletes the custom function from Excel.

#### Syntax
```js
customFunctionObject.delete();
```

#### Parameters
None

#### Returns
void
