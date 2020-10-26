---
title: PowerPoint JavaScript preview APIs
description: 'Details about upcoming PowerPoint JavaScript APIs.'
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Normal
---

# PowerPoint JavaScript preview APIs

New PowerPoint JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Insert and Delete Slides | Allows the insertion of existing slides into the current presentation from another presentation, as well as the ability to delete sildes. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## API list

The following table lists the PowerPoint JavaScript APIs currently in preview. For a complete list of all PowerPoint JavaScript APIs (including preview APIs and previously released APIs), see [all PowerPoint JavaScript APIs](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Specifies which formatting to use during slide insertion.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Specifies the slides from the source presentation that will be inserted into the current presentation. These slides are represented by their IDs which can be retrieved from a `Slide` object.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Specifies where in the presentation the new slides will be inserted.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Inserts the specified slides from a presentation into the current presentation.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Returns an ordered collection of slides in the presentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Deletes the slide from the presentation. Does nothing if the slide does not exist.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Gets the unique ID of the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Gets the number of slides in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Gets a slide using its unique ID. An exception is thrown if the slide does not exist.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Gets a slide using its unique ID. Returns an object whose `isNullObject` property is set to `true` if the slide does not exist.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Gets the loaded child items in this collection.|

## See also

- [PowerPoint JavaScript API Reference Documentation](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API requirement sets](powerpoint-api-requirement-sets.md)
