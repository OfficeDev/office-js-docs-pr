---
title: Adjust worksheet display settings
description: How to adjust some worksheet display settings to make reports easier to read.
ms.date: 04/03/2025
ms.localizationpriority: medium
---

# Adjust worksheet display settings

Excel is often used for reporting scenarios where you want to share worksheet data with others. Your Office Add-in can reduce visual clutter and help focus attention by controlling the appearance of the worksheet. The Office JavaScript API supports changing several visual aspects of the worksheet.

## Page layout and print settings

Add-ins have access to page layout settings at a worksheet level. These control how the sheet is printed. A `Worksheet` object has three layout-related properties: `horizontalPageBreaks`, `verticalPageBreaks`, and `pageLayout`.

`Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` are [PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection) objects. These are collections of [PageBreak](/javascript/api/excel/excel.pagebreak) objects, which specify ranges where manual page breaks are inserted. The following code sample adds a horizontal page break before row **21**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break precedes this range.
    await context.sync();
});
```

`Worksheet.pageLayout` is a [PageLayout](/javascript/api/excel/excel.pagelayout) object. This object contains layout and print settings that aren't dependent on any printer-specific implementation. These settings include margins, orientation, page numbering, title rows, and print area.
The following code sample centers the page (both vertically and horizontally), sets a title row to be printed at the top of every page, and sets the printed area to a subsection of the worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    await context.sync();
});
```

## Turn data type icons on or off (preview)

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Data types can display an icon next to the value in the cell. When you have large tables with many data types, the icons may add visual clutter.

:::image type="content" source="../images/data-types-icon-table.png" alt-text="An Excel table with three data types showing the same icon next to each data type.":::

Use the [Worksheet.showDataTypeIcons](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showdatatypeicons-member) property to toggle data type icons on or off. For more information about data types and their icons, see [Overview of data types in Excel add-ins](excel-data-types-overview.md). The `showDataTypeIcons` property performs the same action as the user toggling data type icons by using the **View** > **Data Type Icons** checkbox. The visibility settings for data type icons are saved with the worksheet and are seen by anyone co-authoring at the time they are changed.

The following code sample shows how to turn off data type icons on a worksheet.

```js
await Excel.run(async (context) => { 
    const sheet = context.workbook.worksheets.getActiveWorksheet(); 
    sheet.showDataTypeIcons = false; 
    await context.sync(); 
});  
```

> [!NOTE]
> If a linked data type displays a **?** icon, this can’t be toggled on or off. Excel needs the user to disambiguate the cell value to find the correct data type. For more information, see [Excel data types: Stocks and geography](https://support.microsoft.com/office/61a33056-9935-484f-8ac8-f1a89e210877).

## Show or hide the worksheet gridlines (preview)

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Gridlines are the faint lines that appear between cells on a worksheet. These can be distracting if you use shapes, icons, or have specific line and border formats on data.

:::image type="content" source="../images/excel-gridlines.png" alt-text="An infographic where the gridlines are distracting.":::

Turn the gridlines on or off with the [Worksheet.showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member) property. This is the same as using the **View** > **Gridlines** checkbox in the Excel UI. The visibility settings for gridlines are saved with the worksheet and are seen by anyone co-authoring at the time they are changed.

The following example shows how to turn off gridlines on a worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.showGridlines = false;
    await context.sync();
});  
```

## Toggle headings (preview)

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Headings are the Excel row numbers that appear on the left side of the worksheet (1, 2, 3) and the column letters that appear at the top of the worksheet (A, B, C). The user may not want these in their report.

:::image type="content" source="../images/excel-heading-label.png" alt-text="A spreadsheet section highlighting the column heading A and the row heading 2.":::

Turn the headings on or off with the [Worksheet.showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member) property. This is the same as using the **View** > **Headings** checkbox in the Excel UI. The following example shows how to turn headings off on a worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.showHeadings = false;
    await context.sync();
});  
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
