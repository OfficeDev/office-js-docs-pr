# Custom Functions Enumerations

Custom functions in Excel uses the following list of enumerations. 

## Values

| **Enumeration**                           | **Value** | **Description**                                                                                                                                                           |
|-------------------------------------------|-----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Excel.CustomFunctionValueType.string      | String         | A string type in Excel.                                                                                                                                                   |
| Excel.CustomFunctionValueType.number      | Number         | A number type in Excel, such as dates, currencies, and other numbers.                                                                                                     |
| Excel.CustomFunctionValueType.invalid      | Invalid         | An invalid type in Excel.                                                                                                     |
| Excel.CustomFunctionValueType.boolean      | Boolean         | A boolean type in Excel.                                                                                                     |
| Excel.CustomFunctionValueType.isodate      | ISODate         | An ISO Date in Excel.                                                                                                     |
| Excel.CustomFunctionDimensionality.scalar | Scalar         | A single cell’s value in Excel.                                                                                                                                            |
| Excel.CustomFunctionDimensionality.matrix | Matrix         | A range of values in Excel, with one or more rows, and one or more columns. In JavaScript, a range is implemented as an array, where each array element is another array. |
