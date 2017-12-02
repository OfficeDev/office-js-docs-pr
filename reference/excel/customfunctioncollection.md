# CustomFunctionCollection Object (JavaScript API for Excel)
A collection of custom functions in Excel. You use this collection to register and unregister custom functions in Excel.

## Methods
| Method | Return Type	|Description| Req. Set|
|:----------------|:-----|:-----------------------------------------------|:----|
|add(name: string)| void | Registers a custom function. The name includes the prefix, period, and function name. For example, CONTOSO.ADD42. | N/A |
|addAll() | void | Deletes all previously registered functions for the add-in, then calls add() on each function registered using Excel.Script.CustomFunctions.| N/A |
|deleteAll() | void | Deletes all previously registered functions for the add-in.| N/A |

## Method Details
### add(name: string)
Registers a custom function. The name includes the prefix, period, and name. For example, CONTOSO.ADD42.

#### Syntax
```js
customFunctionCollectionObject.add(name);
```

#### Parameters
| Parameter	| Type | Description |
|:------------|:---------|:-----------------------------------------------|
| name |	string	| The name of the custom function to add to Excel.|

### Returns
void

### addAll()
Deletes all previously registered functions for the add-in, and then calls add() on each function registered using Excel.Script.CustomFunctions.

#### Syntax
```js
customFunctionCollectionObject.addAll();
```

#### Parameters
None

#### Returns
void

### deleteAll()
Deletes all custom functions added by this add-in.

#### Syntax
```js
customFunctionCollectionObject.deleteAll();
```

#### Parameters
None

#### Returns
void
