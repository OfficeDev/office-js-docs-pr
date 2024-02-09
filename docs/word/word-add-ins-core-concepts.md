---
title: Word JavaScript object model in Office Add-ins
description: Learn about the key components in the Word-specific JavaScript object model.
ms.date: 02/24/2023
ms.topic: concept-article
ms.localizationpriority: high
---

# Word JavaScript object model in Office Add-ins

This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins.

> [!IMPORTANT]
> See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Word APIs and how they work with the document.

## Office.js APIs for Word

A Word add-in interacts with objects in Word by using the Office JavaScript API. This includes two JavaScript object models:

* **Word JavaScript API**: The [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that work with the document, ranges, tables, lists, formatting, and more.

* **Common APIs**: The [Common API](/javascript/api/office) give access to features such as UI, dialogs, and client settings that are common across multiple Office applications.

While you'll likely use the Word JavaScript API to develop the majority of functionality in add-ins that target Word, you'll also use objects in the Common API. For example:

* [Office.Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of document configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the Word application where the add-in is running.
* [Office.Document](/javascript/api/office/office.document): The `Office.Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running. This is separate from the [Word.Document](/javascript/api/word/word.document) object.

![Differences between the Word JS API and Common APIs.](../images/word-js-api-common-api.png)

## Word-specific object model

To understand the Word APIs, you must understand how the components of a document are related to one another.

* The **Document** contains the **Section**s, and document-level entities such as settings and custom XML parts.
* A **Section** contains a **Body**.
* A **Body** gives access to **Paragraph**s, **ContentControl**s, and **Range** objects, among others.
* A **Range** represents a contiguous area of content, including text, white space, **Table**s, and images. It also contains most of the text manipulation methods.
* A **List** represents text in a numbered or bulleted list.

## See also

* [Word JavaScript API overview](../reference/overview/word-add-ins-reference-overview.md)
* [Build your first Word add-in](../quickstarts/word-quickstart.md)
* [Word add-in tutorial](../tutorials/word-tutorial.md)
* [Word JavaScript API reference](/javascript/api/word)
* [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
