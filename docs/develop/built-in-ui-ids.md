---
title: Find the IDs of built-in Office ribbon tabs
description: Find the IDs of the built-in Office ribbon tabs.
ms.date: 07/31/2024
ms.topic: reference
ms.localizationpriority: medium
---

<!-- 
  This article is deliberately left out of the Office Add-ins TOC because 
  it will be moving over to the M365 doc set as soon as that is up and running. 
-->

# Find the IDs of built-in Office ribbon tabs

The following table shows the valid built-in Office ribbon tab `id` values by application. The columns indicate on which platform the IDs are supported. For example, **TabHome** is supported in Word 2016 or later on Windows and Word on the web, but **TabBlogPost** is supported only in Word desktop.

| Office application | Supported on desktop, on the web,</br>and new Outlook on Windows | Supported only on desktop |
|--------------------|--------------------------------------|---------------------------|
| Excel              | TabHome</br>TabInsert</br>TabPageLayoutExcel</br>TabFormulas</br>TabData</br>TabReview</br>TabView</br>TabDeveloper</br>TabAddIns | TabPrintPreview</br>TabBackgroundRemoval |
| OneNote            | TabHome</br>TabInsert</br>TabView | TabDeveloper</br>TabAddIns |
| Outlook            | TabDefault</br>(Depending on what Outlook window is open,</br> this ID refers to either the **Home**, **Message**, or **Meeting** tab.) |                           |
| PowerPoint         | TabHome</br>TabInsert</br>TabDesign</br>TabTransitions</br>TabAnimations</br>TabSlideShow</br>TabReview</br>TabView</br>TabDeveloper</br>TabAddIns | TabPrintPreview</br>TabMerge</br>TabGrayscale</br>TabBlackAndWhite</br>TabBroadcastPresentation</br>TabSlideMaster</br>TabHandoutMaster</br>TabNotesMaster</br>TabBackgroundRemoval</br>TabSlideMasterHome |
| Word               | TabHome</br>TabInsert</br>TabWordDesign</br>TabPageLayoutWord</br>TabReferences</br>TabMailings</br>TabReviewWord</br>TabView</br>TabDeveloper</br>TabAddIns | TabBlogPost</br>TabBlogInsert</br>TabPrintPreview</br>TabOutlining</br>TabConflicts</br>TabBackgroundRemoval</br>TabBroadcastPresentation |
