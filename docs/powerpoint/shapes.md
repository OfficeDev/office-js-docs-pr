---
title: Work with shapes using the PowerPoint JavaScript API
description: 'Learn how to add, remove, and format shapes on PowerPoint slides.'
ms.date: 10/06/2021
ms.localizationpriority: medium
---

# Work with shapes using the PowerPoint JavaScript API (preview)

This article describes how to use geometric shapes, lines, and text boxes in conjunction with the [Shape](/javascript/api/powerpoint/poweroint.shape) and [ShapeCollection](/javascript/api/powerpoint/poweroint.shapecollection) APIs.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## Create shapes

Shapes are created through and stored in a slide's shape collection (`slide.shapes`). `ShapeCollection` has several `.add*` methods for this purpose. All shapes have names and IDs generated for them when they are added to the collection. These are the `name` and `id` properties, respectively. `name` can be set by your add-in.

### Geometric shapes

A geometric shape is created with one of the overrides of `ShapeCollection.addGeometricShape`. The first parameter is either a [GeometricShapeType](/javascript/api/powerpoint/poweroint.geometricshapetype) enum or the string equivalent of one of the enum's values. There is an optional second parameter of type [ShapeAddOptions](/javascript/api/powerpoint/poweroint.shapeaddoptions) that can specify the initial size of the line and its position relative to the top and left sides of the slide, measured in points. Or these properties can be set after the shape is created.

The following code sample creates a rectangle named **"Square"** that is positioned 100 points from the top and left sides of the slide. The method returns a `Shape` object.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
});
```

### Lines

A line is created with one of the overrides of `ShapeCollection.addLine` The first parameter is either a [ConnectorType](/javascript/api/powerpoint/poweroint.connectortype) enum or the string equivalent of one of the enum's values to specify how the line contorts between endpoints. There is an optional second parameter of type [ShapeAddOptions](/javascript/api/powerpoint/poweroint.shapeaddoptions) that can specify the start and end points of the line. Or these properties can be set after the shape is created. The method returns a `Shape` object.

> [!NOTE]
> When the shape is a line, the `top` and `left` properties of the `Shape` and`ShapeAddOptions` objects specify the starting point of the line relative to the top and left edges of the slide. The `height` and `width` properties specify the endpoint of the line *relative to the start point*. So, the end point relative to the top and left edges of the slide is (`top` + `height`) by (`left` + `width`). The unit of measure for all properties is points and negative values are allowed.

The following code sample creates a straight line on the slide.

```js
// This sample creates a straight line on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    return context.sync();
});
```

### Text boxes

A text box is created with the [addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_) method. The first parameter is the text that should appear in the box initially. There is an optional second parameter of type [ShapeAddOptions](/javascript/api/powerpoint/poweroint.shapeaddoptions) that can specify the initial size of the text box and its position relative to the top and left sides of the slide. Or these properties can be set after the shape is created.

The following code sample shows how to create a text box on the first slide.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    return context.sync();
});
```

## Move and resize shapes

Shapes sit on top of the slide. Their placement is defined by the `left` and `top` property. These act as margins from slide's respective edges, measured in points, with [0, 0] being the upper-left corner. The shape size is specified by the `height` and `width` properties. Your code can move or resize the shape by resetting these properties. (These properties have a slightly different meaning when the shape is a line. See [Lines](#lines).

## Text in shapes

Geometric Shapes can contain text. Shapes have a `textFrame` property of type [TextFrame](/javascript/api/powerpoint/poweroint.textframe). The `TextFrame` object manages the text display options (such as margins and text overflow). `TextFrame.textRange` is a [TextRange](/javascript/api/powerpoint/poweroint.textrange) object with the text content and font settings.

The following code sample creates a geometric shape named "Wave" with the text "Shape Text". It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.

```js
// This sample creates a light-blue rectangle with braces ("{}") on the left and right ends and adds the purple text "Shape text" to the center.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var wave = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;
    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");
    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    return context.sync();
});
```

## Delete shapes

Shapes are removed from the slide with the `Shape` object's `delete` method.

The following code sample shows how to delete slides.

```js
PowerPoint.run(function (context) {
    // Delete all shapes from the first slide.
    var sheet = context.presentation.slides.getItemAt(0);
    var shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    // This is required because the items property is being read by the forEach method.
    shapes.load("items/$none");
    return context.sync()
        .then(function () {
            shapes.items.forEach(function (shape) {
                shape.delete()
            });
            return context.sync();
        })
       .catch(errorHandlerFunction);
});
```
