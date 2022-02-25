---
title: Icon element in the manifest file
description: Defines Image elements for Button or Menu controls.
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# Icon element

Defines a set of **Image** elements for [Button](control-button.md) or [Menu](control-menu.md) controls.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) when the parent **VersionOverrides** is type Taskpane 1.0.
- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) when the parent **VersionOverrides** is type Mail 1.0.
- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) when the parent **VersionOverrides** is type Mail 1.1.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  No  | The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`. |

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Image](#image)        | Yes |   resid of an image to use         |

### Image

An image for the button. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> If this image is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.

## Additional requirements for mobile form factors

When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`. This attribute specifies the `UIScreen.scale` property for iOS devices. For more information, see [scale](https://developer.apple.com/documentation/uikit/uiscreen/1617836-scale).

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
