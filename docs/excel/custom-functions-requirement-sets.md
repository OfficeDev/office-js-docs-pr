---
title: Custom Functions requirement sets
description: 'Details about the Custom Functions requirement sets for Excel JavaScript API.'
ms.date: 09/09/2020
ms.prod: excel
localization_priority: Normal
---

# Custom Functions requirement sets

[Custom Functions](./custom-functions-overview.md) use separate requirement sets from the core Excel JavaScript APIs. The following table lists the Custom Functions requirement sets, the supported Office client applications, and the build versions or number for those applications.

|  Requirement set  |  Office on Windows<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.13127.20296 or later | Not supported | 16.40.20081000 or later | July 2020 |
| CustomFunctionsRuntime 1.2 | 16.0.12527.20194 or later | Not supported | 16.34.20020900 or later | January 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 or later | Not supported | 16.34 or later | May 2019 |


> [!NOTE]
> Excel custom functions are not supported on Office 2019 or earlier (one-time purchase).

## CustomFunctionsRuntime 1.1, 1.2, and 1.3

The CustomFunctionsRuntime 1.1 is the first version of the API. Version 1.2 adds the `CustomFunctions.Error` object to support error handling. Version 1.3 adds [XLL streaming](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) support and new `ErrorCode` options to the [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object. 

## See also

- [Custom Functions Reference Documentation](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md)
