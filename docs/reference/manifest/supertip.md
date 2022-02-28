---
title: Supertip element in the manifest file
description: The Supertip element defines a rich tooltip (both title and description).
ms.date: 02/04/2022
ms.localizationpriority: medium
---

# Supertip

Defines a rich tooltip (both Title and Description). It is used by both [Button controls](control-button.md) and [Menu controls](control-menu.md).

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) when the parent **VersionOverrides** is type Taskpane 1.0.
- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) when the parent **VersionOverrides** is type Mail 1.0.
- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) when the parent **VersionOverrides** is type Mail 1.1.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
| [Title](#title) | Yes | The text for the supertip. |
| [Description](#description) | Yes | The description for the supertip.<br>**Note**: (Outlook) Only Windows and Mac clients are supported. |

### Title

Required. The text for the supertip. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

### Description

Required. The description for the supertip. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.

> [!NOTE]
> For Outlook, only Windows and Mac clients support the **Description** element.

## Example

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
