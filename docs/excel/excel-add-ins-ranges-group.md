---
title: Group ranges using the Excel JavaScript API
description: Learn how to group rows or columns of a range together to create an outline using the Excel JavaScript API.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Group ranges for an outline using the Excel JavaScript API

This article provides a code sample that shows how to group ranges for an outline using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Group rows or columns of a range for an outline

Rows or columns of a range can be grouped together to create an [outline](https://support.microsoft.com/office/08ce98c4-0063-4d42-8ac7-8278c49e9aff). These groups can be collapsed and expanded to hide and show the corresponding cells. This makes quick analysis of top-line data easier. Use [Range.group](/javascript/api/excel/excel.range#excel-excel-range-group-member(1)) to make these outline groups.

An outline can have a hierarchy, where smaller groups are nested under larger groups. This allows the outline to be viewed at different levels. Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1)) method. Note that Excel only supports eight levels of outline groups.

The following code sample creates an outline with two levels of groups for both the rows and columns. The subsequent image shows the groupings of that outline. In the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example). A group defines what will be collapsed, not the row or column with the control.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    await context.sync();
});
```

:::image type="content" source="../images/excel-outline.png" alt-text="Range with a two-level, two-dimension outline.":::

## Remove grouping from rows or columns of a range

To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1)) method. This removes the outermost level from the outline. If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
