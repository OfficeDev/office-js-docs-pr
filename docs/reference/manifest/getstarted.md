---
title: GetStarted element in the manifest file
description: Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.
ms.date: 02/22/2022
ms.localizationpriority: medium
---

# GetStarted element

Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md). If the **GetStarted** element is omitted, the callout uses the values from the [DisplayName](displayname.md) and [Description](description.md) elements instead.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

## Child elements

| Element                       | Required | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Yes      | The title used for the top of the callout.     |
| [Description](#description)   | Yes      | The description / body content for the callout.|
| [LearnMoreUrl](#learnmoreurl) | Yes       | A URL to a page that explains the add-in in detail.   |

### Title 

Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.

### Description

Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.

### LearnMoreUrl

Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.

> [!NOTE]
> **LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients. We recommend that you add this URL for all clients so that the URL will render when it becomes available. 

## See also

The following code samples use the **GetStarted** element.

* [Excel Web Add-in for Manipulating Table and Chart Formatting](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Word Add-in JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
