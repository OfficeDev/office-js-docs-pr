---
title: Bind and refresh shapes in your PowerPoint add-in
description: Learn how to create, load, update, and delete shape bindings in a PowerPoint add-in by using the PowerPoint JavaScript API.
ms.topic: how-to
ms.date: 05/21/2026
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Bind to shapes in a PowerPoint presentation

Use shape bindings when your add-in needs to find and update specific shapes without requiring users to select them again. A binding gives a shape a persistent identifier that your add-in can reuse later.

Create a binding by calling [BindingCollection.add](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-add-member(1)) with a unique ID. Then use that ID to get the shape and read or update its properties across sessions.

Shape bindings provide these benefits.

- Establishes a relationship between the add-in and the shape in the document. Bindings are persisted in the document and can be accessed at a later time.
- Enables access to shape properties to read or update, without requiring the user to select any shapes.

## Quick workflow

If you want to try shape bindings quickly, use this flow.

1. Create and bind a shape. See [Create a bound shape in PowerPoint](#create-a-bound-shape-in-powerpoint).
1. Update the bound shape when new image data is available. See [Refresh a bound shape with updated data](#refresh-a-bound-shape-with-updated-data).
1. Reload bindings when your add-in starts. See [Load bindings](#load-bindings).
1. Remove stale bindings when needed. See [Delete a binding](#delete-a-binding).

The following image shows how an add-in might bind to two shapes on a slide. Each shape has a binding ID created by the add-in: `star` and `pie`. Using the binding ID, the add-in can access the desired shape to update properties.

:::image type="content" source="../images/powerpoint-bind-shapes.png" alt-text="Binding to a star shape with the ID 'star' and binding to a pie chart with the ID 'pie'.":::

## Scenario: Use bindings to sync with a data source

A common use case is keeping shapes synced with images from a data source. Without bindings, users often recopy and repaste images whenever the source changes. With bindings, your add-in can fetch the newest image and update the correct shape automatically.

In a general implementation, consider these two components.

1. **The data source**. Any source of data or an asset library, such as Microsoft SharePoint or Microsoft OneDrive.
1. **The PowerPoint add-in**. The add-in gets data from the data source, converts it to a Base64-encoded image, inserts a shape, and binds the shape with a unique identifier. Later, when image data changes, the add-in uses the binding identifier to find and update that same shape.

> [!NOTE]
> You decide the implementation details of how to sync updates from the data source and how to get or create images. This article only describes how to use the Office JS APIs in your add-in to bind a shape and update it with the latest images.

## Create a bound shape in PowerPoint

Use the `PowerPoint.BindingCollection.add()` method for the presentation to create a binding which refers to a particular shape.

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
});
```

Call `BindingCollection.add` to add the binding to the bindings collection in PowerPoint. The following sample shows how to add a new binding for a shape to the bindings collection.

```javascript
// Create a binding ID to track the shape for later updates. 
const bindingId = "productChart"; 
// Create binding by adding the new shape to the bindings collection. 
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
            let myBindings = context.presentation.bindings;
            myBindings.load("items");
            await context.sync();

            // Log all binding IDs to console.
            if (myBindings.items.length > 0) {
                myBindings.items.forEach(async (binding) => {
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

When a shape is deleted, its associated binding is also removed from the PowerPoint binding collection. Any object references you have to the binding, or shape, will return errors if you access any properties or methods on those objects. Be sure to handle potential error scenarios for a deleted shape if your add-in keeps Binding or Shape objects.

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

When maintaining freshness on shapes, you may also want to check the zOrder. See the [zOrderPosition](/javascript/api/powerpoint/powerpoint.shape) property for more information.

- [Work with shapes using the PowerPoint JavaScript API](shapes.md)
- [Bind to regions in a document or spreadsheet](../develop/bind-to-regions-in-a-document-or-spreadsheet.md)
