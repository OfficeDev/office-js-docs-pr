---
title: PowerPoint JavaScript object model in Office Add-ins
description: Learn about the key components in the PowerPoint-specific JavaScript object model.
ms.date: 06/18/2025
ms.topic: concept-article
ms.localizationpriority: high
---

# PowerPoint JavaScript object model in Office Add-ins

This article describes concepts that are fundamental to using the [PowerPoint JavaScript API](../reference/overview/powerpoint-add-ins-reference-overview.md) to build add-ins.

## Office.js APIs for PowerPoint

A PowerPoint add-in interacts with objects in PowerPoint by using the Office JavaScript API. This includes two JavaScript object models:

- **PowerPoint JavaScript API**: The [PowerPoint JavaScript API](/javascript/api/powerpoint) provides strongly-typed objects that work with the presentation, slides, tables, shapes, formatting, and more. To learn about the asynchronous nature of the PowerPoint APIs and how they work with the presentation, see [Using the application-specific API model](../develop/application-specific-api-model.md).

- **Common APIs**: The [Common API](/javascript/api/office) give access to features such as UI, dialogs, and client settings that are common across multiple Office applications. To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

While you'll likely use the PowerPoint JavaScript API to develop the majority of functionality in add-ins that target PowerPoint, you'll also use objects in the Common API. For example:

- [Office.Context](/javascript/api/office/office.context): The `Office.Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of presentation configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the PowerPoint application where the add-in is running.
- [Office.Document](/javascript/api/office/office.document): The `Office.Document` object provides the `getFileAsync()` method, which you can use to download the PowerPoint file where the add-in is running. It also provides the `getActiveViewAsync()` method, which you can use to check whether the presentation is in a "read" or "edit" view. "edit" corresponds to any of the views in which you can edit slides: Normal, Slide Sorter, or Outline View. "read" corresponds to either Slide Show or Reading View.

## PowerPoint-specific object model

To understand the PowerPoint APIs, you must understand how key components of a presentation are related to one another.

- The presentation contains slides and presentation-level entities such as settings and custom XML parts.
- A slide contains content like shapes, text, and tables.
- A layout determines how a slide's content is organized and displayed.

For the full set of objects supported by the PowerPoint JavaScript API, see [PowerPoint JavaScript API](/javascript/api/powerpoint).

## See also

- [PowerPoint JavaScript API overview](../reference/overview/powerpoint-add-ins-reference-overview.md)
- [Build your first PowerPoint add-in](../quickstarts/powerpoint-quickstart-yo.md)
- [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial-yo.md)
- [PowerPoint JavaScript API reference](/javascript/api/powerpoint)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
