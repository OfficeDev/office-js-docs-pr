---
title: PowerPoint JavaScript preview APIs
description: 'Details about upcoming PowerPoint JavaScript APIs.'
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
---

# PowerPoint JavaScript preview APIs

New PowerPoint JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Slide management | Adds support for getting references to slides, slide layouts, and slide masters. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Shapes | Adds support for getting references to the shapes in a slide. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## API list

The following table lists the PowerPoint JavaScript APIs currently in preview. For a complete list of all PowerPoint JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Specifies the ID of a Slide Layout to be used for the new slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Specifies the ID of a Slide Master to be used for the new slide.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Returns the collection of `SlideMaster` objects that are in the presentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Gets the unique ID of the shape.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Gets the number of shapes in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Gets a shape using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Gets a shape using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Gets a shape using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Gets the loaded child items in this collection.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Gets the layout of the slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Returns a collection of shapes in the slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Gets the `SlideMaster` object that represents the slide's default content.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Adds a new slide at the end of the collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Gets the unique ID of the slide layout.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Gets the name of the slide layout.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Gets the number of layouts in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Gets a layout using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Gets a layout using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Gets a layout using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Gets the loaded child items in this collection.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Gets the unique ID of the Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Gets the collection of layouts provided by the Slide Master for slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Gets the unique name of the Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Gets the number of Slide Masters in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Gets a Slide Master using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Gets a Slide Master using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Gets a Slide Master using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Gets the loaded child items in this collection.|
