---
title: Create and manage shapes in Excel add-ins
description: Learn how to create, position, format, group, export, and delete shapes in Excel worksheets by using the Excel JavaScript API.
ms.date: 06/03/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Create, format, and manage shapes with the Excel JavaScript API

Use shapes when your add-in needs to place visual elements on a worksheet, such as callouts, process diagrams, labels, images, or connectors. This article shows how to create geometric shapes, lines, images, and text boxes, then move, resize, group, export, and delete them by using the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.

The following image shows shapes arranged as a thermometer.
:::image type="content" source="../images/excel-shapes.png" alt-text="Image of a thermometer made as an Excel shape.":::

## Create shapes

All of a worksheet's shapes are stored in `Worksheet.shapes`. Use the `ShapeCollection` `add*` methods to create them. When you add a shape, Excel generates both an `id` and a `name`. Your add-in can also set `name` so it can retrieve the shape later with `ShapeCollection.getItem(name)`.

Use the following methods to create common shape types.

| Shape | Add method | Signature |
|-------|------------|-----------|
| Geometric shape | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Image (JPEG or PNG) | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| Text box | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### Create a geometric shape

Use `ShapeCollection.addGeometricShape` when your add-in needs a built-in shape, such as a rectangle, wave, or chevron. The method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum value.

The following example creates a 150-by-150 pixel rectangle named **Square** and places it 100 pixels from the top-left corner of **MyWorksheet**.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    const rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### Add an image or SVG

Use `addImage` when a user selects a local JPEG or PNG file in your task pane. Pass the image as a Base64-encoded string. Use `addSvg` when you already have SVG markup as XML.

The following example loads a local image with [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader). It removes the `base64,` prefix from the data URL and inserts the image as a shape.

```js
const myFile = document.getElementById("selectedFile");
const reader = new FileReader();

reader.onload = () => {
    Excel.run(async (context) => {
        const startIndex = reader.result.toString().indexOf("base64,");
        const myBase64 = reader.result.toString().substring(startIndex + 7);
        const sheet = context.workbook.worksheets.getItem("MyWorksheet");
        const image = sheet.shapes.addImage(myBase64);
        image.name = "Image";

        await context.sync();
    }).catch(errorHandlerFunction);
};

reader.readAsDataURL(myFile.files[0]);
```

### Add a line or connector

Use `ShapeCollection.addLine` to draw a line between two points. Provide the left and top coordinates for the start point and end point. You can also pass a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to control how the line bends between those points.

The following example creates a straight line on the worksheet.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";

    await context.sync();
});
```

Lines can also connect to other shapes. Use `connectBeginShape` and `connectEndShape` to attach the start and end of the line to shape-specific connection points. Use `Shape.connectionSiteCount` to confirm that the connection point index is valid. Use `disconnectBeginShape` and `disconnectEndShape` to remove those connections.

The following example connects the line named **MyLine** to the shapes named **LeftShape** and **RightShape**.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);

    await context.sync();
});
```

## Move and resize shapes

Shapes sit on top of the worksheet grid. The `left` and `top` properties control their position, with `[0, 0]` at the upper-left corner of the worksheet. To adjust an existing position, use `incrementLeft` and `incrementTop`.

To rotate a shape, set the `rotation` property or call `incrementRotation`. To change its depth order relative to other shapes, call `setZOrder` with a [ShapeZOrder](/javascript/api/excel/excel.shapezorder) value.

To resize a shape, set `height` or `width` directly, or use `scaleHeight` and `scaleWidth`. The scaling methods take a [ShapeScaleType](/javascript/api/excel/excel.shapescaletype) value and can also take a [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) value to control which point stays anchored. If `lockAspectRatio` is `true`, scaling keeps the current aspect ratio.

> [!NOTE]
> Setting `height` or `width` directly changes only that dimension, even when `lockAspectRatio` is `true`.

The following example rotates the shape named **Octagon** by 30 degrees and scales it to 125% of its current width while keeping the top-left corner fixed.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("MyWorksheet");
    const shape = sheet.shapes.getItem("Octagon");

    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft
    );

    await context.sync();
});
```

## Work with the active shape

Use one of the following methods when your add-in needs to work with the shape the user currently selected.

- [Workbook.getActiveShape](/javascript/api/excel/excel.workbook) when a selected shape is required
- [Workbook.getActiveShapeOrNullObject](/javascript/api/excel/excel.workbook) when your code should continue even if no shape is selected

The following example gets the active shape, checks whether one is selected, and increases its height by 10%.

```js
await Excel.run(async (context) => {
    const shape = context.workbook.getActiveShapeOrNullObject();
    shape.load(["isNullObject", "height"]);

    await context.sync();

    if (!shape.isNullObject) {
        shape.height = shape.height * 1.1;
        await context.sync();
    }
});
```

## Add text to shapes

Geometric shapes can contain text. Use the `textFrame` property to access a [TextFrame](/javascript/api/excel/excel.textframe). The `TextFrame` object controls layout details such as margins and text overflow. Use `TextFrame.textRange` to access the [TextRange](/javascript/api/excel/excel.textrange) object that contains the text and font settings.

The following example creates a wave shape, adds centered text, and applies fill and font colors.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
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

Use `ShapeCollection.addTextBox` when you want the same kind of text box that users create from **Insert** > **Text Box** in Excel. The method creates a rectangular shape with a white background and black text, and it accepts the initial text as a string.

The following example creates a small text box that displays **Hello!**.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";

    await context.sync();
});
```

## Group shapes

Group shapes when users should move, resize, or format several shapes as one object. A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is itself a `Shape`, so your add-in can work with the group as a single item.

The following example groups three existing shapes and then moves the group 50 pixels to the right.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const square = shapes.getItem("Square");
    const pentagon = shapes.getItem("Pentagon");
    const octagon = shapes.getItem("Octagon");

    const shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";

    await context.sync();
});

await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);

    await context.sync();
});
```

> [!IMPORTANT]
> After shapes are grouped, access the individual members through `ShapeGroup.shapes`, which is a [GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection). They're no longer available through the worksheet's `shapes` collection. For example, if a worksheet had three shapes and you grouped all three, `shapes.getCount` returns `1`.

## Export shapes as images

Use [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) when your add-in needs to save a shape or reuse it elsewhere. The method returns a Base64-encoded string. Pass a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum value to choose the output format.

The following example exports the shape named **Image** as a PNG.

```js
await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    const shape = shapes.getItem("Image");
    const stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in can use the Base64 string to save the image
    // as a file or insert it into HTML.
});
```

## Delete shapes

Use `Shape.delete()` to remove a shape from the worksheet. If you need to clear many shapes, load the collection first and then delete each item.

The following example deletes every shape from **MyWorksheet**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("MyWorksheet");
    const shapes = sheet.shapes;

    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach((shape) => {
        shape.delete();
    });

    await context.sync();
});
```

## See also

- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
- [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md)
- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)
