---
title: Word JavaScript API overview
description: Overview of the Word JavaScript API.
ms.date: 05/29/2025
ms.topic: concept-article
ms.service: word
ms.localizationpriority: high
---

# Word JavaScript API overview

A Word add-in interacts with objects in Word by using the Office JavaScript API, which includes two JavaScript object models:

* **Word JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Word. Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.

* **Common APIs**: The [Common API](/javascript/api/office), introduced with Office 2013, can be used to access features such as UI, dialogs, and client settings that are common across multiple Office applications.

This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web, or Word 2016 and later. For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).

## Learn programming concepts

See [Word JavaScript object model in Office Add-ins](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.

## Learn about API capabilities

Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md), [use search options in your Word add-in to find text](../../word/search-option-guidance.md), and more. See the table of contents for the complete list of available articles.

For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).

For detailed information about the Word JavaScript API object model, see the [Word JavaScript API reference documentation](/javascript/api/word).

## Try out code samples in Script Lab

Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API. You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.

## See also

* [Word add-ins documentation](../../word/index.yml)
* [Word add-ins overview](../../word/word-add-ins-programming-overview.md)
* [Word JavaScript API reference](/javascript/api/word)
* [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
