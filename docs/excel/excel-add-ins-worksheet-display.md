---
title: Control worksheet display settings
description: Learn how to use the Excel JavaScript API to control page layout, data type icons, gridlines, and headings in a worksheet.
ms.date: 06/03/2026
ms.topic: how-to
ai-usage: ai-assisted
ms.localizationpriority: medium
---

# Control worksheet display settings with the Excel JavaScript API

Use worksheet display settings to make dashboards, reports, and printouts easier to read. By using the Excel JavaScript API, your add-in can control page layout, page breaks, data type icons, gridlines, and headings so users see the worksheet the way you intend.

## Key points

- Use `Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` to control manual page breaks.
- Use `Worksheet.pageLayout` to control print settings such as centering, title rows, and print area.
- Use preview worksheet properties to show or hide data type icons, gridlines, and headings.
- Changes to worksheet display settings are saved with the worksheet.

## Configure page layout and print settings

Use worksheet page layout settings when your add-in prepares a worksheet for printing or sharing. These settings help you control where pages break and which parts of the worksheet print.

### Add a manual page break

The `Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` properties return [PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection) objects. Each collection contains [PageBreak](/javascript/api/excel/excel.pagebreak) objects that define where Excel inserts manual page breaks.

The following code sample adds a horizontal page break before row **21**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break precedes this range.
    await context.sync();
});
```

### Set print layout options

The `Worksheet.pageLayout` property returns a [PageLayout](/javascript/api/excel/excel.pagelayout) object. Use it to control print settings that don't depend on a specific printer, such as margins, orientation, page numbering, title rows, and print area.

The following code sample centers the page, repeats the first row at the top of every printed page, and limits printing to range **A1:D100**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the printed area to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    await context.sync();
});
```

## Control worksheet visuals

Use the following properties to reduce visual clutter in a worksheet before users review or present it.

### Show or hide data type icons

Data types can display an icon next to the value in a cell. In large tables, those icons can distract from the data.

:::image type="content" source="../images/data-types-icon-table.png" alt-text="An Excel table with three data types showing the same icon next to each data type.":::

Use the [Worksheet.showDataTypeIcons](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showdatatypeicons-member) property to show or hide data type icons. This property is equivalent to the user selecting **View** > **Data Type Icons**. The setting is saved with the worksheet and is visible to coauthors when it changes. For more information about data types and their icons, see [Overview of data types in Excel add-ins](excel-data-types-overview.md).

If a linked data type shows a **?** icon, you can't toggle that icon. Excel requires the user to disambiguate the cell value first. For more information, see [Excel data types: Stocks and geography](https://support.microsoft.com/office/61a33056-9935-484f-8ac8-f1a89e210877).

The following code sample hides data type icons on the active worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.showDataTypeIcons = false;
    await context.sync();
});
```

### Show or hide gridlines

Gridlines are the faint lines between cells on a worksheet. They can distract from shapes, icons, or custom borders in a report.

:::image type="content" source="../images/excel-gridlines.png" alt-text="An infographic where the gridlines are distracting.":::

Use the [Worksheet.showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member) property to show or hide gridlines. This property is equivalent to the user selecting **View** > **Gridlines**. The setting is saved with the worksheet and is visible to coauthors when it changes.

The following code sample hides gridlines on the active worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.showGridlines = false;
    await context.sync();
});
```

### Show or hide headings

Headings are the row numbers on the left side of the worksheet and the column letters across the top. In a polished report view, you might want to hide them.

:::image type="content" source="../images/excel-heading-label.png" alt-text="A spreadsheet section highlighting the column heading A and the row heading 2.":::

Use the [Worksheet.showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member) property to show or hide headings. This property is equivalent to the user selecting **View** > **Headings**.

The following code sample hides headings on the active worksheet.

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
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
