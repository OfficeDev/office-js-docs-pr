---
title: Work with Shapes using the Excel JavaScript API
description: ''
ms.date: 2/13/2019
localization_priority: Normal
---

# Work with Shapes using the Excel JavaScript API (preview)

> [!NOTE]
> The APIs discussed in this article are currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

Excel defines shapes as any object that sits on the drawing layer of Excel. That means anything outside of a cell is a shape. This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/shape.md) and [ShapeCollection](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/shapecollection.md) APIs.

## Create shapes

Shapes are created by adding new shapes to a worksheet's ShapeCollection object. This is done through the `add*` methods. The following shapes are currently supported in this manner:

- Geometric Shape - `addGeometricShape`.
- Image (either JPEG or PNG) - `addImage`.
- Line - `addLine`.
- SVG - `addSvg`.
- Text Box - `addTextBox`.

### Geometric shapes

### Images and SVGs

### Lines

### Text boxes

## Move and resize shapes

## Delete shapes

## Text in shapes

## Shape groups

## See also

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)
- [Work with Charts using the Excel JavaScript API](excel-add-ins-charts.md)
