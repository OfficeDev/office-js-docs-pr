---
title: Common coding issues and unexpected platform behaviors
description: 'A list of Office JavaScript API platform issues frequently encountered by developers.'
ms.date: 01/02/2020
localization_priority: Normal
---

# Common coding issues and unexpected platform behaviors

This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome. If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.

## Common APIs and Outlook APIs are not promise-based

The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model. Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes. For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).

These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes. If you need `await` behavior, you can wrap the method call in an explicitly created Promise.

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).

## Some properties must be set with JSON structs

> [!NOTE]
> This section only applies to the host-specific APIs for Excel and Word.

Some properties must be set as JSON structs, instead of setting their individual subproperties. One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout). The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`. That statement throws an error because `zoom` is not loaded. Even if `zoom` were to be loaded, the set of scale will not take effect. All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.

This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format). Properties of `format` can be set using object navigation, as shown here:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier. All read-only properties can have their non-read-only subproperties directly set. Writeable properties like `PageLayout.zoom` must be set with a JSON struct. In summary:

- Read-only property: Subproperties can be set through navigation.
- Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).

## Excel data transfer limits

If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:

- Excel on the web has a payload size limit for requests and responses of 5MB. `RichAPI.Error` will be thrown if that limit is exceeded.
- A range is limited to five million cells for get operations.

If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`. Split the operation into smaller pieces as needed. Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.

These limitations are typically exceeded by large ranges. Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range. See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.

## Setting read-only properties

The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only. If you attempt to set a read-only property, the write operation will fail silently, with no error thrown. The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## Removing event handlers

Event handlers must be removed using the same `RequestContext` in which they were added. If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## See also

- [OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.
- [Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs. Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.
- [UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.
