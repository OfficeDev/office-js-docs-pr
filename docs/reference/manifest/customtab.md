---
title: CustomTab element in the manifest file
description: On the ribbon, you specify which tab and group for their add-in commands.
ms.date: 09/02/2021
ms.localizationpriority: medium
---

# CustomTab element

On the ribbon, specify the tab and group for your add-in commands. This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.

On custom tabs, the add-in can have custom or built-in groups. Add-ins are limited to one custom tab.

The **id** attribute must be unique within the manifest.

> [!IMPORTANT]
> In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Group](group.md)      | No |  Defines a Group of commands.  |
|  [OfficeGroup](#officegroup)      | No |  Represents a built-in Office control group. **Important**: Not available in Outlook. |
|  [Label](#label-tab)      | Yes |  The label for the CustomTab or a Group.  |
|  [InsertAfter](#insertafter)      | No |  Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Only available in PowerPoint. |
|  [InsertBefore](#insertbefore)      | No |  Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Only available in PowerPoint. |

### Group

Optional, but if not present there must be at least one **OfficeGroup** element. See [Group element](group.md). The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.

### OfficeGroup

Optional, but if not present there must be at least one **Group** element. Represents a built-in Office control group. The **id** attribute specifies the ID of the built-in Office group. To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.

> [!IMPORTANT]
> The `OfficeGroup` element is not available in Outlook.

### Label (Tab)

Required. The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

### InsertAfter

Optional. Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview". (See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element. You cannot have both **InsertAfter** and **InsertBefore**.

> [!IMPORTANT]
> The `InsertAfter` element is only available in PowerPoint.

### InsertBefore

Optional. Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview". (See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element. You cannot have both **InsertAfter** and **InsertBefore**.

> [!IMPORTANT]
> The `InsertBefore` element is only available in PowerPoint.
