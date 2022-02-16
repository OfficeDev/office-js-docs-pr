---
title: Custom Functions requirement sets
description: 'Details about the Custom Functions requirement sets for Excel JavaScript API.'
ms.date: 02/15/2022
ms.prod: excel
ms.localizationpriority: medium
---

# Custom Functions requirement sets

[Custom Functions](../../excel/custom-functions-overview.md) use separate requirement sets from the core Excel JavaScript APIs. The following table lists the Custom Functions requirement sets, the supported Office client applications, and the build versions or number for those applications.

|  Requirement set  |  Office 2021 or later on Windows<br>(one-time purchase)  |  Office on Windows<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2021 and later)  | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.14326.20454 or later | 16.0.13127.20296 or later | Not supported | 16.40.20081000 or later | July 2020 |
| CustomFunctionsRuntime 1.2 | 16.0.14326.20454 or later | 16.0.12527.20194 or later | Not supported | 16.34.20020900 or later | January 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.14326.20454 or later | 16.0.12527.20092 or later | Not supported | 16.34 or later | May 2019 |

## CustomFunctionsRuntime 1.1, 1.2, and 1.3

The CustomFunctionsRuntime 1.1 is the first version of the API. Requirement set 1.2 adds the `CustomFunctions.Error` object to support error handling. Requirement set 1.3 adds [XLL streaming](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) support and new `ErrorCode` options to the [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object.

## See also

- [Custom Functions Reference Documentation](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
