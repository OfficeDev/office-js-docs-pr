---
title: Bind to shapes in a PowerPoint presentation
description: Learn how to bind shapes and access them from your add-in to keep them up-to-date.
ms.topic: how-to
ms.date: 03/31/2025
ms.localizationpriority: medium
---

# Bind to shapes in a PowerPoint presentation

Your PowerPoint add-in can bind to shapes to consistently access them through an identifier. Your add-in establishes a binding by calling `BindingCollection.add` and assigning a unique identifier. The identifier can be used at any time to reference the shape and access its properties. Creating bindings provides the following value to your add-in.

- Establishes a relationship between the add-in and the shape in the document. Bindings are persisted in the document and can be accessed at a later time.
- Enables access to shape properties to read or update, without requiring the user to select any shapes.

The following image shows how an add-in might bind to two shapes on a slide. Each shape has a binding ID created by the add-in: `star` and `pie`. Using the binding ID the add-in can access the desired shape to update properties.

:::image type="content" source="../images/powerpoint-bind-shapes.png" alt-text="Diagram of binding to a star shape with the ID 'star' and binding to a pie chart with the ID 'pie'":::

## Scenario: Use bindings to sync with a data source

A common scenario for using bindings is to help users keep shapes up to date with a data source. Often when a user creates a presentation, they copy and paste images from the data source into the presentation. Over time, to keep the images up to date, they manually copy and paste the latest images from the data source. Your add-in can help automate this process by retrieving up-to-date images from the data source on the user’s behalf. When the user needs a shape fill updated, your add-in can use the binding to find the correct shape, and update the shape fill with the newer image.

In a general implementation, there are two components to consider for binding a shape in PowerPoint and updating it with a new image from a data source.

1. **The data source**. This is any source of data or asset library such as Microsoft SharePoint or Microsoft OneDrive.  
1. **The PowerPoint add-in**. The add-in gets data from the data source, based on what the user needs. It converts the data to a base64 formatted image. This is the only fill type the bound shape can accept. It inserts a shape upon the user’s request and binds it with a unique identifier. Then it fills the shape with the base64 image based on the original data source. Shapes are updated upon the user’s request and the add-in uses the binding identifier to find the shape and update the image with the last saved base64 image.

> [!NOTE]
> The implementation details of syncing to updates from the data source and how to get images, or create images are for you to decide. This article only describes how to use the Office JS APIs in your add-in to bind a shape and update it with latest images.

## Create a bound shape in PowerPoint

Use the `PowerPoint.BindingsCollection.add()` method for the presentation to create a binding which refers to a particular shape.

:::image type="content" source="../images/powerpoint-steps-to-bind-shape.png" alt-text="Diagram showing add-in creating base64 image from data source, then creating the shape from the image, and adding a unique ID.":::

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

    // Fill shape with a base64 image. 
    // Note: The image is typically created from a data source request. 
    const productsImage = "...base64 image data...";
    myShape.fill.setImage(productsImage);
});
```

Call `BindingsCollection.add` to add the binding to the bindings collection in PowerPoint. The following sample shows how to add a new binding for a shape to the bindings collection.

```javascript
//Create a binding ID to track the shape for later updates. 
const bindingId = "productChart"; 
// Create binding by adding the new shape to the bindings collection. 
context.presentation.bindings.add(myShape, PowerPoint.BindingType.shape, bindingId); 
```

## Refresh a bound shape with updated data

When there is an update to the image data you can refresh the shape image by finding it via the binding identifier. The following code sample shows how to find a bound shape with the identifier and fill it with an updated image. The updated image is created by the add-in based on the data source request, or provided by the data source directly.

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
        context.presentation.bindings.getItemAt(bindingId).delete();
        await context.sync();
    });
}
```

## Load bindings

When a presentation is opened and your add-in first loads, you can load all the bindings to continue working with them. The following code shows how to load all bindings in a presentation and display them in the console.

```javascript
async function loadBindings() {
    await PowerPoint.run(async (context) => {
        try {
            let myBindings = context.presentation.bindings;
            myBindings.load("items");
            await context.sync();

            // Log all binding ids to console.
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
            const binding = context.presentation.bindings.getItemAt(id);
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

When maintaining freshness on shapes, you may also want to check the zOrder. See the [zOrderPosition](/javascript/api/powerpoint/powerpoint.shape?view=powerpoint-js-preview) property for more information.

- [Work with shapes using the PowerPoint JavaScript API](shapes.md)
- [Bind to regions in a document or spreadsheet](../develop/bind-to-regions-in-a-document-or-spreadsheet.md)
