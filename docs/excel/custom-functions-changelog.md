---
ms.date: 01/04/2019
description: Discover the latest updates to custom functions.
title: Custom Functions Changelog
---

# Custom functions changelog (preview)

Excel custom functions is still in preview and that means there are frequent changes to the product, including changes and the release of new features. This changelog provides the most up-to-date information about any changes to the product.

- **Nov 7, 2017**: Shipped* the custom functions preview and samples
- **Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later
- **Nov 28, 2017**: Shipped* support for cancellation on asynchronous functions (requires change for streaming functions)
- **May 7, 2018**: Shipped* support for Mac, Excel Online, and synchronous functions running in-process
- **September 20, 2018**: Shipped support for custom functions JavaScript runtime. For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).
- **October 20, 2018**: With the [October Insiders build](https://support.office.com/en-us/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online. On Mac, this parameter should be ignored.
- **December 12, 2018**: Custom functions now include a way to discover a cell's address. For more information, see [Discovering cells that invoke custom functions](custom-functions-overview.md#discovering-cells-that-invoke-custom-functions).
- **January 8, 2019**: Binding method `CustomFunctionMapping()` has been altered to `CustomFunctions.associate()`. For more information, see [Custom functions best practices (preview)](custom-functions-best-practices.md)

\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")

For a list of known issues with the product, see [Known Issues](custom-functions-overview.md#known-issues). 

## See also

* [Custom functions overview](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
