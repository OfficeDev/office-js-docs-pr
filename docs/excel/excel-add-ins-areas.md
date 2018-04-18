---
title: Work with multiple ranges simultaneously in Excel Add-ins
description: ''
ms.date: 4/20/2018
---

# Work with multiple ranges simultaneously in Excel Add-ins (Preview)

The Excel JavaScript Library provides APIs to enable your add-in to perform operations and set properties on multiple ranges simultaneously. The ranges do not have to be contiguous with each other. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.

## Areas

The `Range` object has an `areas` property of type `AreaCollection`. The members of the collection are other `Range` objects. (There is no `Area` type.) When there is more than one range in the collection, the parent range object represents the set of all the ranges in its `areas` collection. 

> [!NOTE]
> Excel JavaScript APIs that are in preview are documented in [this branch of the office-js-docs repo](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel). See, for example, [AreaCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel/areacollection.md).

Setting a property on the parent range has the effect of setting the corresponding property on all the child ranges. *Reading* a property of the parent range will usually return `null`, because the corresponding properties of the ranges in its `areas` collection may have inconsistent values for those properties. This will happen even in the case where they all have the same value for the property that is being read. 

> [!NOTE]
> There is always at least one range in the `areas` collection for any `Range` object, but when there is only one, the parent range behaves like an ordinary `Range` object, so you can read as well as write its properties. The `areas` collection has no purpose for this kind of `Range` object. 

The `areas` property is read only. The only way to get a range with multiple child ranges in its `areas` collection is to call an API that returns such a range. The primary API that does this is the [Worksheet.getRange](https://dev.office.com/reference/add-ins/excel/worksheet#getrangeaddress-string) method, which can now take more than one range address as its parameter. Your code passes the addresses as a comma-delimited string.

The following is an example of setting a property on multiple ranges. Note the following about this code:

- It is intended to highlight all ranges with calculated values (that is, with formulas). 
- It assumes that the ranges **F3:F5** and **H3:H5** (and no others) have formulas on the active sheet. Note that these two ranges are *not* contiguous.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("F3:F5, H3:H5");
    range.format.fill.color = "pink";
    
    return context.sync();
})
```

This example applies to scenarios in which you can hard code the range addresses that you pass to `getRange`. This would include the following scenarios, among others:

- The code runs in the context of a known template.
- The code runs in the context of imported data where the schema of the data is known.

If your scenario requires you to discover at runtime which range addresses to pass to `getRange`, then you will need to create one or more helper methods to find the ranges. We are working on a variety of new APIs that will simplify this process, such as an API that will find all cells matching specified properties or value types. We'll update this article when they become available for preview.


## See also

- [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
- [Range Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/range)
