---
title: Common coding issues and unexpected platform behaviors
description: ''
ms.date: 10/24/2019
localization_priority: Normal
---

# Common coding issues and unexpected platform behaviors

This articles highlights subtle discrepancies and unexpected behaviors with the Office JavaScript API. If you encounter an issue that belongs in this list, please let us know by using the GitHub feedback form at the bottom of the article.

## Some properties cannot be set as navigational properties

> [!NOTE]
> This section only applies to the host-specific APIs for Excel and Word.

Some properties must be set as JSON structs, instead of setting their individual subproperties. One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout). The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };

// The following code does throws an error because `zoom` is not loaded.
sheet.pageLayout.zoom.scale = 200;
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

The [TypeScript definitions](https://github.com/DefinitelyTyped/DefinitelyTyped/) for Office-JS specify which object properties are read-only. However, JavaScript developers are able to write code ignoring the `readonly` modifier. In that case, Office-JS silently ignores the write operation. The following example shows the read-only property [Chart.id](/javascript/api/excel/excel.chart#id) erroneously attempting to be set.

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## Tracking event handlers

> [!NOTE]
> This section only applies to the host-specific APIs for Excel, OneNote, and Word.

Event handlers are tied to individual, client-side, proxy objects. The event handler will attach to the related workbook object when synced. Removing an event handler requires a reference to the original proxy object.

```js
Excel.run(function (context) {
    const pieChart = context.workbook.worksheets.getActiveWorksheet().charts.getItem("Pie");
    pieChart.onActivated.add(chartActivated);
    return context.sync().then (function() {
        // This following code will not remove the event handler.
        // It is using a different proxy object for the chart.
        var sameChart = context.workbook.worksheets.getActiveWorksheet().charts.getItem("Pie");
        sameChart.onActivated.remove(chartActivated);
        return context.sync();
    });
});
```

## See also

- [The office-js GitHub repo](https://github.com/OfficeDev/office-js/issues): The issues page is a complete list of known product issues.
- [Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): A question and answer site for professional and enthusiast programmers. Use the "office-js" tag when asking a Stack Overflow question so the community can find it and help.
- [User Voice](https://officespdev.uservoice.com/): A place to suggest new features for the Office platform.
