# Custom Function Enumerations

The following list of enumerations are used by custom functions in Excel.

## Values

| **Enumeration**                           | **Value** | **Description**                                                                                                                                                           |
|-------------------------------------------|-----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Excel.CustomFunctionValueType.string      | 0         | A string type in Excel.                                                                                                                                                   |
| Excel.CustomFunctionValueType.number      | 1         | A number type in Excel, such as dates, currencies, and other numbers.                                                                                                     |
| Excel.CustomFunctionDimensionality.scalar | 0         | A single cell’s value in Excel                                                                                                                                            |
| Excel.CustomFunctionDimensionality.matrix | 1         | A range of values in Excel, with one or more rows, and one or more columns. In JavaScript, a range is implemented as an array, where each array element is another array. |
