---
title: PowerPoint JavaScript object model in Office Add-ins
description: Learn about the key components in the PowerPoint-specific JavaScript object model.
ms.date: 07/10/2026
ms.topic: concept-article
ms.localizationpriority: high
---

# PowerPoint JavaScript object model in Office Add-ins

The [PowerPoint JavaScript API](/javascript/api/powerpoint) lets your add-in read and modify presentations, slides, shapes, text, and tables. This article introduces the object models you work with and the key objects you use to build PowerPoint add-ins.

## Office.js APIs for PowerPoint

A PowerPoint add-in interacts with objects in PowerPoint by using the Office JavaScript API. This includes two JavaScript object models:

- **PowerPoint JavaScript API**: The [PowerPoint JavaScript API](/javascript/api/powerpoint) provides strongly typed objects that work with the presentation, slides, tables, shapes, formatting, and more. To learn about the asynchronous nature of the PowerPoint APIs and how they work with the presentation, see [Using the application-specific API model](../develop/application-specific-api-model.md).

- **Common APIs**: The [Common API](/javascript/api/office) gives access to features such as UI, dialogs, and client settings that are common across multiple Office applications. To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

While you'll likely use the PowerPoint JavaScript API to develop the majority of functionality in add-ins that target PowerPoint, you'll also use objects in the Common API. For example:

- [Office.Context](/javascript/api/office/office.context): The `Office.Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of presentation configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the PowerPoint application where the add-in is running.
- [Office.Document](/javascript/api/office/office.document): The `Office.Document` object provides the `getFileAsync()` method, which you can use to download the PowerPoint file where the add-in is running. It also provides the `getActiveViewAsync()` method, which you can use to check whether the presentation is in a "read" or "edit" view. "edit" corresponds to any of the views in which you can edit slides: Normal, Slide Sorter, or Outline View. "read" corresponds to either Slide Show or Reading View.

## PowerPoint-specific object model

To understand the PowerPoint APIs, you must understand how key components of a presentation are related to one another.

- The presentation contains slides and presentation-level entities such as settings and custom XML parts.
- A slide contains content like shapes, text, and tables.
- A layout determines how a slide's content is organized and displayed.

The following table lists key objects in the PowerPoint JavaScript API.

| Object | Description |
|--------|-------------|
| [Presentation](/javascript/api/powerpoint/powerpoint.presentation) | The top-level object. Provides access to slides, tags, and presentation-wide settings. |
| [Slide](/javascript/api/powerpoint/powerpoint.slide) | A single slide. Contains shapes and references its layout. |
| [SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection) | The collection of slides in the presentation. |
| [Shape](/javascript/api/powerpoint/powerpoint.shape) | A shape on a slide, such as a text box, image, line, or geometric shape. |
| [TextRange](/javascript/api/powerpoint/powerpoint.textrange) | The text within a shape. Use it to read and edit text. |
| [SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout) | The layout that determines how a slide's content is arranged. |
| [SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster) | The master that defines default formatting and the available layouts. |

For the full set of objects supported by the PowerPoint JavaScript API, see [PowerPoint JavaScript API](/javascript/api/powerpoint).

## See also

- [PowerPoint JavaScript API reference](/javascript/api/powerpoint)
- [PowerPoint add-ins](powerpoint-add-ins.md)
- [Build your first PowerPoint add-in](../quickstarts/powerpoint-quickstart-yo.md)
- [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial-yo.md)
- [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
