---
title: Group ranges using the Excel JavaScript API
description: Learn how to group rows or columns of a range together to create an outline using the Excel JavaScript API.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Group ranges for an outline using the Excel JavaScript API

This article provides code samples that show how to group and ungroup ranges for an outline using the Excel JavaScript API. Grouping rows or columns creates collapsible sections in your worksheet, making it easier to organize and present complex data. This is especially useful for financial reports, hierarchical data, and large datasets where users need to focus on summary information while having details available on demand.

For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Key points

- Use `Range.group` to group rows or columns into collapsible outline sections.
- Use `Range.ungroup` to remove grouping from rows or columns.
- Outlines support up to eight levels of hierarchy for nested groups.
- Use `Worksheet.showOutlineLevels` to programmatically expand or collapse outline levels.
- Grouped ranges don't include the control row or column; only the content that is collapsed.
- Groups can be nested to create multi-level hierarchies for complex data organization.

## Group rows or columns of a range for an outline

Rows or columns of a range can be grouped together to create an [outline](https://support.microsoft.com/office/08ce98c4-0063-4d42-8ac7-8278c49e9aff). These groups can be collapsed and expanded to hide and show the corresponding cells. This makes quick analysis of top-line data easier. Use [Range.group](/javascript/api/excel/excel.range#excel-excel-range-group-member(1)) to create these outline groups.

An outline can have a hierarchy, where smaller groups are nested under larger groups. This allows the outline to be viewed at different levels. Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1)) method. Excel supports up to eight levels of outline groups.

When you group a range, Excel adds outline controls (the plus and minus buttons) outside the grouped range. By default, the control appears on the row or column after the grouped range. For example, if you group rows 4-9, the control appears on row 10. When users click the minus button, rows 4-9 collapse; when they click the plus button, those rows expand again.

### Create a multi-level outline

The following code sample creates an outline with two levels of groups for both the rows and columns. The subsequent image shows the groupings of that outline. The grouped ranges don't include the row or column of the outline control (the "Totals" for this example). A group defines what is collapsed, not the row or column with the control.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // are on row 10, meaning 4-9 collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // are on rows 6 and 9, meaning 4-5 and 7-8 collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // are on column R, meaning C-Q collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // are on columns G, L, and R, meaning C-F, H-K, and M-P collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    await context.sync();
});
```

:::image type="content" source="../images/excel-outline.png" alt-text="Range with a two-level, two-dimension outline.":::

### Group rows for a simple outline

For simpler scenarios, create a single-level outline to organize related data. This example groups quarterly data under an annual summary.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Group rows 3-6 (Q1-Q4 data) so they can be collapsed.
    // The outline control appears on row 7 (Annual Total).
    sheet.getRange("3:6").group(Excel.GroupOption.byRows);

    await context.sync();
});
```

## Control outline visibility levels

After creating a multi-level outline, programmatically expand or collapse specific levels using `Worksheet.showOutlineLevels`. This is useful for presenting data at different detail levels.

The following code sample collapses all groups to show only the highest-level summary data.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Show only level 1 for both rows and columns (most collapsed view).
    // This hides all grouped details and shows only top-level summaries.
    sheet.showOutlineLevels(1, 1);

    await context.sync();
});
```

## Remove grouping from rows or columns of a range

To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1)) method. This removes the outermost level from the outline. If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.

## See also

- [Outline (group) data in a worksheet](https://support.microsoft.com/office/08ce98c4-0063-4d42-8ac7-8278c49e9aff)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
- [Get a range using the Excel JavaScript API](excel-add-ins-ranges-get.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
