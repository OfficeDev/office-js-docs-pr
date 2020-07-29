---
title: Coding guidance for common issues and unexpected platform behaviors
description: 'A list of Office JavaScript API platform issues frequently encountered by developers.'
ms.date: 07/29/2020
localization_priority: Normal
---

# Coding guidance for common issues and unexpected platform behaviors

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

## Some properties cannot be set directly

> [!NOTE]
> This section only applies to the host-specific APIs for Excel and Word.

Some properties cannot be set, despite being writable. These properties are part of a parent property that must be set as a single object. This is because that parent property relies on the subproperties having specific, logical relationships. These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties. One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout). The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`. That statement throws an error because `zoom` is not loaded. Even if `zoom` were to be loaded, the set of scale will not take effect. All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.

This behavior differs from [navigational properties](host-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format). Properties of `format` can be set using object navigation, as shown here:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

You can identify a property that cannot have its subproperties directly set by checking its read-only modifier. All read-only properties can have their non-read-only subproperties directly set. Writeable properties like `PageLayout.zoom` must be set with an object at that level. In summary:

- Read-only property: Subproperties can be set through navigation.
- Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).

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

## Supporting Internet Explorer

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## Excel-specific issues

### API limitations when the active workbook switches

Add-ins for Excel are intended to operate on a single workbook at a time. Errors can arise when a workbook that is separate from the one running the add-in gains focus. This only happens when particular methods are in the process of being called when the focus changes.

The following APIs are affected by this workbook switch:

|Excel JavaScript API | Error thrown |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> This only applies to multiple Excel workbooks open on Windows or Mac.

## See also

- [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md)
- [OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.
- [Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs. Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.
- [UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.
