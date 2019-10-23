---
title: Common coding issues and platform quirks
description: ''
ms.date: 10/23/2019
localization_priority: Normal
---

# Common coding issues and platform quirks

TODO: Add intro

## Some properties cannot be treated as navigational properties

Some properties must be set as JSON structs, instead of setting their individual subproperties. One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout). The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };

// The following code does throws an error because `zoom` is not loaded.
// sheet.pageLayout.zoom.scale = 200;
// Note that even if `zoom` is loaded, the set of scale will not take effect.
// All context operations will happen on `zoom`, refreshing the proxy object in the add-in.
```

This behavior differs from navigational properties like [Range.format](/javascript/api/excel/excel.range#format). Properties of `format` can be set using object navigation, as shown here:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier. All read-only properties can have their non-read-only subproperties directly set. Properties like `PageLayout.zoom` are not read-only and must be set with a JSON struct. In summary:

- Read-only property: Subproperties can be set through navigation.
- Writable property: Subproperties cannot be set through navigation.

## Setting read-only properties

The [TypeScript definitions](https://github.com/DefinitelyTyped/DefinitelyTyped/) for Office-JS specify which object properties are read-only. However, JavaScript developers are able to write code ignoring the `readonly` field. In that case, Office-JS silently ignores the write operation. The following example shows the read-only property [Chart.id](/javascript/api/excel/excel.chart#id) erroneously attempting to be set.

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```
