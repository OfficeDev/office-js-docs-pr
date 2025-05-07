---
title: Work with shapes using the PowerPoint JavaScript API
description: Learn how to add, remove, and format shapes on PowerPoint slides.
ms.date: 05/06/2025
ms.localizationpriority: medium
---

# Work with shapes using the PowerPoint JavaScript API

This article describes how to use geometric shapes, lines, and text boxes in conjunction with the [Shape](/javascript/api/powerpoint/powerpoint.shape) and [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) APIs.

## Create shapes

Shapes are created through and stored in a slide's shape collection (`slide.shapes`). `ShapeCollection` has several `.add*` methods for this purpose. All shapes have names and IDs generated for them when they are added to the collection. These are the `name` and `id` properties, respectively. `name` can be set by your add-in.

### Geometric shapes

A geometric shape is created with one of the overloads of `ShapeCollection.addGeometricShape`. The first parameter is either a [GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) enum or the string equivalent of one of the enum's values. There is an optional second parameter of type [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) that can specify the initial size of the shape and its position relative to the top and left sides of the slide, measured in points. Or these properties can be set after the shape is created.

The following code sample creates a rectangle named **"Square"** that is positioned 100 points from the top and left sides of the slide. The method returns a `Shape` object.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    await context.sync();
});
```

### Lines

A line is created with one of the overloads of `ShapeCollection.addLine`. The first parameter is either a [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) enum or the string equivalent of one of the enum's values to specify how the line contorts between endpoints. There is an optional second parameter of type [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) that can specify the start and end points of the line. Or these properties can be set after the shape is created. The method returns a `Shape` object.

> [!NOTE]
> When the shape is a line, the `top` and `left` properties of the `Shape` and `ShapeAddOptions` objects specify the starting point of the line relative to the top and left edges of the slide. The `height` and `width` properties specify the endpoint of the line *relative to the start point*. So, the end point relative to the top and left edges of the slide is (`top` + `height`) by (`left` + `width`). The unit of measure for all properties is points and negative values are allowed.

The following code sample creates a straight line on the slide.

```js
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(PowerPoint.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    await context.sync();
});
```

### Text boxes

A text box is created with the [addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) method. The first parameter is the text that should appear in the box initially. There is an optional second parameter of type [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) that can specify the initial size of the text box and its position relative to the top and left sides of the slide. Or these properties can be set after the shape is created.

The following code sample shows how to create a text box on the first slide.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    await context.sync();
});
```

## Move and resize shapes

Shapes sit on top of the slide. Their placement is defined by the `left` and `top` properties. These act as margins from slide's respective edges, measured in points, with `left: 0` and `top: 0` being the upper-left corner. The shape size is specified by the `height` and `width` properties. Your code can move or resize the shape by resetting these properties. (These properties have a slightly different meaning when the shape is a line. See [Lines](#lines).)

## Text in shapes

Geometric shapes can contain text. Shapes have a `textFrame` property of type [TextFrame](/javascript/api/powerpoint/powerpoint.textframe). The `TextFrame` object manages the text display options (such as margins and text overflow). `TextFrame.textRange` is a [TextRange](/javascript/api/powerpoint/powerpoint.textrange) object with the text content and font settings.

The following code sample creates a geometric shape named **"Braces"** with the text **"Shape text"**. It also adjusts the shape and text colors, as well as sets the text's vertical alignment to the center.

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    await context.sync();
});
```

## Group and ungroup shapes

In PowerPoint, you can group several shapes and treat them like a single shape. You can subsequently ungroup grouped shapes. To learn more about grouping objects in the PowerPoint UI, see [Group or ungroup shapes, pictures, or other objects](https://support.microsoft.com/office/a7374c35-20fe-4e0a-9637-7de7d844724b).

### Group shapes

To group shapes with the JavaScript API, use [ShapeCollection.addGroup](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgroup-member(1)).

The following code sample shows how to group existing shapes of type [GeometricShape](/javascript/api/powerpoint/powerpoint.shapetype) found on the current slide.

```typescript
// Groups the geometric shapes on the current slide.
await PowerPoint.run(async (context) => {
    // Get the shapes on the current slide.
    context.presentation.load("slides");
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("shapes/items/type,shapes/items/id");
    await context.sync();

    const shapes = slide.shapes;
    const shapesToGroup = shapes.items.filter((item) => item.type === PowerPoint.ShapeType.geometricShape);
    if (shapesToGroup.length === 0) {
        console.warn("No shapes on the current slide, so nothing to group.");
        return;
    }

    // Group the geometric shapes.
    console.log(`Number of shapes to group: ${shapesToGroup.length}`);
    const group = shapes.addGroup(shapesToGroup);
    group.load("id");
    await context.sync();

    console.log(`Grouped shapes. Group ID: ${group.id}`);
});
```

### Ungroup shapes

To ungroup shapes with the JavaScript API, get the [group](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-group-member) property from the group's `Shape` object then call [ShapeGroup.ungroup](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-ungroup-member(1)).

The following code sample shows how to ungroup the first shape group found on the current slide.

```js
// Ungroups the first shape group on the current slide.
await PowerPoint.run(async (context) => {
    // Get the shapes on the current slide.
    context.presentation.load("slides");
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("shapes/items/type,shapes/items/id");
    await context.sync();

    const shapes = slide.shapes;
    const shapeGroups = shapes.items.filter((item) => item.type === PowerPoint.ShapeType.group);
    if (shapeGroups.length == 0) {
        console.warn("No shape groups on the current slide, so nothing to ungroup.");
        return;
    }

    // Ungroup the first grouped shapes.
    const firstGroupId = shapeGroups[0].id;
    const shapeGroupToUngroup = shapes.getItem(firstGroupId);
    shapeGroupToUngroup.group.ungroup();
    await context.sync();

    console.log(`Ungrouped shapes with group ID: ${firstGroupId}`);
});
```

## Delete shapes

Shapes are removed from the slide with the `Shape` object's `delete` method.

The following code sample shows how to delete shapes.

```js
await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const shapes = context.presentation.slides.getItemAt(0).shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
        
    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    await context.sync();
});
```

## See also

- [Work with tables using the PowerPoint JavaScript API](work-with-tables.md)
- [Bind to shapes in a PowerPoint presentation](bind-shapes-in-presentation.md)
- [Group or ungroup shapes, pictures, or other objects](https://support.microsoft.com/office/a7374c35-20fe-4e0a-9637-7de7d844724b)
