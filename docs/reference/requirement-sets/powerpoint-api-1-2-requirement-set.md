---
title: PowerPoint JavaScript API requirement set 1.2
description: 'Details about the PowerPointApi 1.2 requirement set.'
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# What's new in PowerPoint JavaScript API 1.2

PowerPointApi 1.2 added support for inserting slides from another presentation into the current presentation and for deleting slides.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Insert and Delete Slides](../../powerpoint/insert-slides-into-presentation.md) | Allows the insertion of existing slides into the current presentation from another presentation, as well as the ability to delete slides. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## API list

The following table lists the PowerPoint JavaScript API requirement set 1.2. For a complete list of all PowerPoint JavaScript APIs (including preview APIs and previously released APIs), see [all PowerPoint JavaScript APIs](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Specifies which formatting to use during slide insertion.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|Specifies the slides from the source presentation that will be inserted into the current presentation.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|Specifies where in the presentation the new slides will be inserted.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|Inserts the specified slides from a presentation into the current presentation.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Returns an ordered collection of slides in the presentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|Deletes the slide from the presentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Gets the unique ID of the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|Gets the number of slides in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|Gets a slide using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Gets the loaded child items in this collection.|

## See also

- [PowerPoint JavaScript API Reference Documentation](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [PowerPoint JavaScript API requirement sets](powerpoint-api-requirement-sets.md)
