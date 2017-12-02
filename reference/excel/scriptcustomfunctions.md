# Script.CustomFunctions Object (JavaScript API for Excel)
Use the Script.CustomFunctions object to define metadata about a custom function before registration.

## Properties
| Property| Type| Description	| Req. Set |
|:---------------|:--------|:----------------|:----------|
| call | string	| The name of the function that contains the JavaScript code of the custom function.	| N/A |
| description | string | Description of the function for the autocomplete menu. | N/A |
| helpUrl	| string	| URL of the help page for the function. | N/A |
| result	| object	| Result returned by the function.	| N/A|
| result.resultType | [Excel.CustomFunctionValueType](customfunctionsenumerations.md) | Result type returned by the function.	| N/A |
| result.resultDimensionality | [Excel.CustomFunctionDimensionality](customfunctionsenumerations.md) | The dimensionality of the result returned by the function.	| N/A |
| parameters | array| The parameters of the function.	| N/A |
| parameters.name | string	| Name of a parameter passed to the custom function.	| N/A |
| parameters.description	| string	| Description of the parameter.| N/A |
| parameters.valueType	| [Excel.CustomFunctionValueType](customfunctionsenumerations.md) | The type of the parameter.| N/A |
| parameters.valueDimensionality	| [Excel.CustomFunctionDimensionality](customfunctionsenumerations.md) | The dimensionality of the parameter.	| N/A |
| Options	| object | Specifies options to change the behavior of the custom function. The complete list of options is described below.	| N/A |

### Custom function options
The following Boolean properties are used to change the behavior of a custom function.

| Property | Description |
|:---------------|:--------|
| batch	| Sets the custom function to run in batch mode. The custom function takes an array, where each array element is an array of parameters. The default option is false when the batch option is not specified. |
| stream | Sets the custom function to run in streamed mode. The default option is false when the stream option is not specified. |

#### Example
```js
Excel.Script.customFunctions["CONTOSO"]["ADD42"] = {
    call: add42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [
        {
            name: "num 1",
            description: "The first number",
            valueType: Excel.CustomFunctionValueType.number,
            valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        {
            name: "num 2",
            description: "The second number",
            valueType: Excel.CustomFunctionValueType.number,
            valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
        }
    ],
    options:{ batch: false, stream: false }
};
```
