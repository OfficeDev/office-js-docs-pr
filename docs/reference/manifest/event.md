---
title: Event element in the manifest file
description: Defines an event handler in an add-in.
ms.date: 01/11/2022
ms.localizationpriority: medium
---

# Event element

Defines an event handler in an add-in.

> [!NOTE]
> For information about support and usage, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Type](#type-attribute)  |  Yes  | Specifies the event to handle. |
|  [FunctionExecution](#functionexecution-attribute)  |  Yes  | Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported. |
|  [FunctionName](#functionname-attribute)  |  Yes  | Specifies the function name for the event handler. |

### Type attribute

Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.

|  Event type  |  Description  |
|:-----|:-----|
|  `ItemSend`  |  The event handler will be invoked when the user sends a message or meeting invitation.  |

### FunctionExecution attribute

Required. MUST be set to `synchronous`.

### FunctionName attribute

Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
