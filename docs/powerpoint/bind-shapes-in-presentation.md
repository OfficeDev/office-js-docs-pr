---
title: Bind and refresh shapes in PowerPoint add-ins
description: Learn how to bind PowerPoint shapes to stable IDs so your add-in can refresh, load, and delete shape references reliably.
ms.topic: how-to
ms.date: 05/29/2026
ms.localizationpriority: medium
---

# Bind and refresh shapes in a PowerPoint presentation

Use shape bindings when your add-in needs to find and update the same shape later, such as refreshing a number from an external data source.

A binding creates a stable identifier for a shape. Your add-in can use that identifier to get the shape again, update it, and handle cases where the shape was deleted.

Bindings provide two key benefits:

- They establish a relationship between the add-in and the shape in the document. The document persists bindings, so you can access them later.
- They enable access to shape properties for reading or updating, without requiring the user to select any shapes.

The following image shows how an add-in might bind to two shapes on a slide. Each shape has a binding ID created by the add-in: `star` and `pie`. By using the binding ID, the add-in can access the desired shape to update its properties.

:::image type="content" source="../images/powerpoint-bind-shapes.png" alt-text="Binding to a star shape with the ID 'star' and binding to a pie chart with the ID 'pie'.":::

## Scenario: Sync shapes with a data source

A common scenario is keeping presentation visuals up to date from a data source. Instead of manually replacing images, an add-in can retrieve the latest data, convert it to a Base64 image, and update the correct shape by using its binding ID.

In a general implementation, consider two components for binding a shape in PowerPoint and updating it with a new image from a data source.

1. **The data source**. This is any source of data or asset library such as Microsoft SharePoint or Microsoft OneDrive.
1. **The PowerPoint add-in**. The add-in gets data from the data source based on what the user needs. It converts the data to a Base64-encoded image. This is the only fill type the bound shape can accept. It inserts a shape upon the user’s request and binds it with a unique identifier. Then it fills the shape with the Base64 image based on the original data source. Shapes are updated upon the user’s request and the add-in uses the binding identifier to find the shape and update the image with the last saved Base64 image.

> [!NOTE]
> You decide the implementation details for syncing updates and creating images. This article focuses on using the Office.js APIs to bind shapes and refresh them.

## Create a bound shape in PowerPoint

Use [BindingCollection.add](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-add-member(1)) to create a binding that refers to a specific shape.

:::image type="content" source="../images/powerpoint-steps-to-bind-shape.png" alt-text="Add-in creates a Base64-encoded image from data source, then creates the shape from the image and adds a unique ID.":::

The following sample shows how to create a shape on the first selected slide.

```javascript
await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();

    // Insert new shape on first selected slide.
    const myShape = slides
        .getItemAt(0)
        .shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
            top: 100,
            left: 30,
            width: 200,
            height: 200
        });

    // Fill shape with a Base64-encoded image.
    // Note: The image is typically created from a data source request.
    const productsImage = "...Base64 image data...";
    myShape.fill.setImage(productsImage);

    await context.sync();
});
```

Call `BindingCollection.add` to add the shape to the PowerPoint bindings collection. The following sample continues from the previous sample and adds a new binding for `myShape`.

```javascript
// Create a binding ID to track the shape for later updates.
const bindingId = "productChart";
// Create a binding by adding the shape to the bindings collection.
context.presentation.bindings.add(myShape, PowerPoint.BindingType.shape, bindingId);
```

## Refresh a bound shape with updated data

After there's an update to the image data, refresh the shape image by finding it via the binding identifier. The following code sample shows how to find a bound shape with the identifier and fill it with an updated image. The image is updated by the add-in based on the data source request or provided by the data source directly.

```javascript
async function updateBinding(bindingId, image) {
    await PowerPoint.run(async (context) => {
        try {
            // Get the shape based on binding ID.
            const myShape = context.presentation.bindings
                .getItem(bindingId)
                .getShape();

            // Update the shape to latest image.
            myShape.fill.setImage(image);
            await context.sync();

        } catch (err) {
            console.error(err);
        }
    });
}
```

## Delete a binding

The following sample shows how to delete a binding by deleting it from the bindings collection.

```javascript
async function deleteBinding(bindingId) {
    await PowerPoint.run(async (context) => {
        context.presentation.bindings.getItem(bindingId).delete();
        await context.sync();
    });
}
```

## Load bindings

When a user opens a presentation and your add-in first loads, you can load all the bindings to continue working with them. The following code shows how to load all bindings in a presentation and display them in the console.

```javascript
async function loadBindings() {
    await PowerPoint.run(async (context) => {
        try {
            const myBindings = context.presentation.bindings;
            myBindings.load("items");
            await context.sync();

            // Log all binding IDs to console.
            if (myBindings.items.length > 0) {
                myBindings.items.forEach((binding) => {
                    console.log(binding.id);
                });
            }
        } catch (err) {
            console.error(err);
        }
    });
}
```

## Error handling when a binding or shape is deleted

When you delete a shape, PowerPoint also removes its associated binding from the binding collection. Any references to that binding or shape can fail when you call methods or access properties. Handle these error scenarios if your add-in stores `Binding` or `Shape` objects.

The following code shows one approach to error handling when a binding object references a deleted binding. Use a try/catch statement and then call a function to reload all binding and shape references when an error occurs.

```javascript
async function getShapeFromBindingID(id) {
    await PowerPoint.run(async (context) => {
        try {
            const binding = context.presentation.bindings.getItem(id);
            const shape = binding.getShape();

            await context.sync();
            return shape;
        } catch (err) {
            console.log(err);
            return undefined;
        }
    });
}
```

## See also

When maintaining visual freshness, you might also want to check shape layering by using [Shape.zOrderPosition](/javascript/api/powerpoint/powerpoint.shape).

- [Create and format shapes in PowerPoint add-ins](shapes.md)
- [PowerPoint JavaScript object model in Office Add-ins](core-concepts.md)
- [Build your first PowerPoint task pane add-in](../quickstarts/powerpoint-quickstart-yo.md)
- [Tutorial: Create a PowerPoint task pane add-in](../tutorials/powerpoint-tutorial-yo.md)
- [Bind to regions in a document or spreadsheet](../develop/bind-to-regions-in-a-document-or-spreadsheet.md)
- [PowerPoint JavaScript API reference](/javascript/api/powerpoint)
