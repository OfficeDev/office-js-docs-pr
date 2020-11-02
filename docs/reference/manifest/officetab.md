---
title: OfficeTab element in the manifest file
description: The OfficeTab element defines the ribbon tab where your add-in command appears.
ms.date: 06/20/2019
localization_priority: Normal
---

# OfficeTab element

Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in. This element is required.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  Group      | Yes |  Defines a group of commands. You can add only one group per add-in to the default tab.  |

The following are valid tab `id` values by application. Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).

### Outlook

- **TabDefault**

### Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval

### PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## Group

A group of UI extension points in a tab. A group can have up to six controls. The **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).

## OfficeTab example

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
