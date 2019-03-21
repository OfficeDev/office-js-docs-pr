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

A geometric shape is created with `ShapeCollection.addGeometricShape`. That method takes a [GeometricShapeType](//javascript/api/excel/excel.geometricshapeyype) enum as an argument.

The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.

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
    return context.sync();
}).catch(errorHandlerFunction);
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
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

Any Shape object can be converted to an image. [Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.
> The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum.

### Lines

A line is created with `ShapeCollection.addLine`. That method needs the left and top margins of the line's start and end points. It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints. The following code sample creates a straight line on the worksheet.

```js
// This creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

Lines can be connected to other Shape objects. The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points. The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds. A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.

The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.

```js
// This connects a line between two shapes at connection points '0' and '3'.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    return context.sync();
}).catch(errorHandlerFunction);
```

## Move and resize shapes

Shapes sit on top of the worksheet. Their placement is defined by the `left` and `top` property. These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner. These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods. How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.

A shape's depth relative to other shapes is the `zorderPosition` property. This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder` adjusts the ordering of the current shape relative to the other shapes.

Your add-in has a couple options for changing the height and width of shapes. Setting the `height` and `width` properties change that dimension without changing the other dimension. The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)). An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner). If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.

> [!NOTE]
> Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.

The following code sample shows a shape being scaled to 1.25 times its original and rotated 30 degrees.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);
    return context.sync();
}).catch(errorHandlerFunction);
```

## Text in shapes

Geometric Shapes can display text.

## Shape groups

## Delete shapes

Shapes are removed from the worksheet with the `Shape` object's `delete` method.

```js
// This deletes all the shapes from "MyWorksheet".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
            shape.delete()
        });
        return context.sync();
    }).catch(errorHandlerFunction);
}).catch(errorHandlerFunction);
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)
- [Work with Charts using the Excel JavaScript API](excel-add-ins-charts.md)
