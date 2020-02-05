---
title: Call Excel APIs from a custom function
description: Learn which Excel APIs you can call from your custom function. 
ms.date: 02/05/2020
localization_priority: Normal
---
# Call Excel APIs from a custom function

Custom functions are able to call most Office.js Excel API to get range data and obtain more context for your calculations. 

[TODO: Add disclaimer]

Calling these APIs through a custom function can be helpful if:

- A custom function needs to get information from Excel before calculation. This information might include document properties, range formats, custom XML parts, a workbook name, or any other number of Excel-specific information.

- A custom function will set the cell's number format for the return values after calculation.

## Code sample

The code sample below shows you how to do TODO. This sample will only work if you have added changes to your manifest and your task pane's HTML file as show in the [Share data and events between Excel custom functions and task pane tutorial](TODO LINK).

```js
TODO
```

## Restrictions to calling Office.js through a custom function

Calling Office.js through a custom function is not recommended if:

- The API you intend to call **writes** or **sets** information. This can result in proor performance, time outs, and infinite loops.

## Next steps

## See also
[Share data and events between Excel custom functions and task pane tutorial](TODO link).