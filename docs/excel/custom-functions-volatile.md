---
ms.date: 04/28/2019
description: Learn to implement volatile and offline streaming custom functions.
title: Volatile values in functions (preview)
localization_priority: Normal
---

## Volatile values in functions

Volatile functions are functions in which the value changes each time the cell is calculated. The value can change even if none of the function's arguments change. These functions recalculate every time Excel recalculates. For example, imagine a cell that calls the function `NOW`. Every time `NOW` is called, it will automatically return the current date and time.

Excel contains several built-in volatile functions, such as `RAND` and `TODAY`. For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling. For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.

If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
