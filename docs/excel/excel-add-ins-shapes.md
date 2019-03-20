---
title: Work with Shapes using the Excel JavaScript API
description: ''
ms.date: 03/14/2019
localization_priority: Normal
---

# Work with Shapes using the Excel JavaScript API (preview)

> [!NOTE]
> The APIs discussed in this article are currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

Excel defines shapes as any object that sits on the drawing layer of Excel. That means anything outside of a cell is a shape. This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape]/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs. [Charts](/javascript/api/excel/excel.chart) are covered in [their own article](excel-add-ins-charts.md)).

## Create shapes

Shapes are created by adding new shapes to a worksheet's ShapeCollection object. This is done through the `add*` methods. Created shapes are stored in the worksheet's `ShapeCollection` (`Worksheet.shapes`). Shape can have a relevant name stored in the `name` property, which is then accessible through the `ShapeCollection.getItem(name)` method.

The following shapes are added using the associated method:

| Shape | Add Method | Signature |
|-------|------------|-----------|
| Geometric Shape | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Image (either JPEG or PNG) | [addImage](/javascript/api/excel/excel.shapecollection?view=office-js#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection?view=office-js#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection?view=office-js#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| Text Box | [addTextBox](/javascript/api/excel/excel.shapecollection?view=office-js#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### Geometric shapes

A geometric shape is created with `ShapeCollection.addGeometricShape`. That method takes a [GeometricShapeType](//javascript/api/excel/excel.geometricshapeyype) enum as an argument. The following code creates a hexagon shape and names it **"Hexagon"**.

```js
// This creates a hexagon shape named "Hexagon".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.hexagon);
    shape.name = "Hexagon";
    return context.sync();
});
```

### Images

JPEG, PNG, and SVG images can be inserted into a worksheet as shapes. The `ShapeCollection.addImage` method takes a base64-encoded string as an argument. This is either a JPEG or PNG image in string form. `ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.

The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string. The string has the metadata "base64," removed before the shape is created.

```js
// This creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = event.target.result.indexOf("base64,");
        var myBase64 = event.target.result.substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    });
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

> [!TIP]
> Any Shape object can be converted to an image. [Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.
> The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum.

### Lines

A line is created with `ShapeCollection.addLine`. That method needs the left and top margins of the line's start and end points. It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.

```js
// This creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
});
```

Lines can be connected to other Shape objects. The `connectBeginShape` and `connectEndShape` 

### Text boxes

## Move and resize shapes

```js
// This creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
```

## Delete shapes

## Text in shapes

## Shape groups

## See also

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)
- [Work with Charts using the Excel JavaScript API](excel-add-ins-charts.md)
