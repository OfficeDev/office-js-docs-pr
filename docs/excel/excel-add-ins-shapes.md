---
title: Work with shapes using the Excel JavaScript API
description: Learn how Excel defines shapes as any object that sits on the drawing layer of Excel.
ms.date: 04/14/2025
ms.localizationpriority: medium
---

# Work with shapes using the Excel JavaScript API

Excel defines shapes as any object that sits on the drawing layer of Excel. That means anything outside of a cell is a shape. This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs. [Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).

The following image shows shapes which form a thermometer.
:::image type="content" source="../images/excel-shapes.png" alt-text="Image of a thermometer made as an Excel shape.":::

## Create shapes

Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`). `ShapeCollection` has several `.add*` methods for this purpose. All shapes have names and IDs generated for them when they are added to the collection. These are the `name` and `id` properties, respectively. `name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.

The following types of shapes are added using the associated method.

| Shape | Add Method | Signature |
|-------|------------|-----------|
| Geometric Shape | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Image (either JPEG or PNG) | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| Text Box | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### Geometric shapes

A geometric shape is created with `ShapeCollection.addGeometricShape`. That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.

The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### Images

JPEG, PNG, and SVG images can be inserted into a worksheet as shapes. The `ShapeCollection.addImage` method takes a Base64-encoded string as an argument. This is either a JPEG or PNG image in string form. `ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.

The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string. The string has the metadata "base64," removed before the shape is created.

```js
// This sample creates an image as a Shape object in the worksheet.
let myFile = document.getElementById("selectedFile");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        let startIndex = reader.result.toString().indexOf("base64,");
        let myBase64 = reader.result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getItem("MyWorksheet");
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### Lines

A line is created with `ShapeCollection.addLine`. That method needs the left and top margins of the line's start and end points. It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints. The following code sample creates a straight line on the worksheet.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

Lines can be connected to other Shape objects. The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points. The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds. A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.

The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    await context.sync();
});
```

## Move and resize shapes

Shapes sit on top of the worksheet. Their placement is defined by the `left` and `top` property. These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner. These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods. How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.

A shape's depth relative to other shapes is defined by the `zorderPosition` property. This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder` adjusts the ordering of the current shape relative to the other shapes.

Your add-in has a couple options for changing the height and width of shapes. Setting either the `height` or `width` property changes the specified dimension without changing the other dimension. The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)). An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner). If the `lockAspectRatio` property is `true`, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.

> [!NOTE]
> Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.

The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");

    let shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);

    await context.sync();
});
```

## Get the active shape

Get the active shape by using either of the following methods.

- [Workbook.getActiveShape](/javascript/api/excel/excel.workbook)
- [Workbook.getActiveShapeOrNullObject](/javascript/api/excel/excel.workbook)

The following code sample shows how to get the active shape in a workbook and increase its height by 10%.

```javascript
 await Excel.run(async (context) => {
    const shape = context.workbook.getActiveShapeOrNullObject();
    if (shape !== null) {
      shape.load("height");
      await context.sync();
      shape.height = shape.height + shape.height * 0.1;
      await context.sync();
    }
  });
```

## Text in shapes

Geometric Shapes can contain text. Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe). The `TextFrame` object manages the text display options (such as margins and text overflow). `TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.

The following code sample creates a geometric shape named "Wave" with the text "Shape Text". It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;

    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");

    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;

    await context.sync();
});
```

The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text. This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.

The following code sample shows the creation of a text box with the text "Hello!".

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    await context.sync();
});
```

## Shape groups

Shapes can be grouped together. This allows a user to treat them as a single entity for positioning, sizing, and other related tasks. A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.

The following code sample shows three shapes being grouped together. The subsequent code sample shows that shape group being moved to the right 50 pixels.

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let square = shapes.getItem("Square");
    let pentagon = shapes.getItem("Pentagon");
    let octagon = shapes.getItem("Octagon");

    let shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    await context.sync();
});

// This sample moves the previously created shape group to the right by 50 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    await context.sync();
});
```

> [!IMPORTANT]
> Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection). They are no longer accessible through the worksheet's shape collection after being grouped. As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.

## Export shapes as images

Any `Shape` object can be converted to an image. [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) returns Base64-encoded string. The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.

```js
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shape = shapes.getItem("Image");
    let stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in may use the Base64-encoded string to save the image as a file or insert it in HTML.
});
```

## Delete shapes

Shapes are removed from the worksheet with the `Shape` object's `delete` method. No other metadata is needed.

The following code sample deletes all the shapes from **MyWorksheet**.

```js
// This deletes all the shapes from "MyWorksheet".
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");
    let shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    
    await context.sync();
});
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)
- [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md)
