---
title: Requirements element in the manifest file
description: The Requirements element specifies the minimum requirement set and methods your Office Add-in needs to be activated by Office or to override base manifest settings.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# Requirements element

The meaning of this element depends on whether it is used [in the base manifest](#in-the-base-manifest)] or [as a child of a **VersionOverrides** element](#as-a-child-of-a-versionoverrides-element).

> [!TIP]
> Before using this element, be familiar with [Specify Office hosts and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)

## In the base manifest

When used in the base manifest, that is, as a direct child of [OfficeApp](officeapp.md), the **Requirements** element specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that your Office Add-in needs to be activated by Office. The add-in will not be activated on any combination of Office version and platform (such as Windows, Mac, iOS, and web) that does not support the specified methods and requirement sets.

**Add-in type:** Task pane, Mail

## As a child of a VersionOverrides element

When used as a child of [VersionOverrides](versionoverrides.md), specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that must be supported by the Office version and platform (such as Windows, Mac, iOS, and web) in order for the settings in the **VersionOverrides** element *that override base manifest settings* to take effect. If the requirements are not met, but the requirements specified in the base manifest *are* met (or there is no **Requirements** element in the base manifest), then the add-in will be activated, but the settings in the base manifest will be used, not the overriding settings in the **VersionOverrides**. However, child elements in the **VersionOverrides** *that support additional features, rather than override base manifest settings,* will still be applied unless:

1. They are associated with a requirement set.
2. There is a **Requirements** element in the **VersionOverrides** that specifies that requirement set.
3. The current platform and Office version combination does not support the requirement set.

Examples of elements in a **VersionOverrides** that provide additional features instead of overriding settings in the base manifest include **WebApplicationInfo** and **EquivalentAddins**.

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

### Remarks

The **Requirements** element serves no purpose in a **VersionOverrides** if it specifies no additional requirements that aren't specfied in a **Requirements** in the base manifest, because if the Office version and platform don't support the requirements in the base manifest, the add-in is not activated and the **VersionOverrides** element is not even parsed. For this reason, you should use a **Requirements** element in a **VersionOverrides** only when both of these conditions are met:

- Your add-in has extra features that are implemented with configuration in a **VersionOverrides** (such as Add-in Commands), and that require a method or requirement set that is *not* specified in a **Requirements** element in the base manifest.
- Your add-in is useful, and should be activated (but without the extra features), even in a combination of platform and Office version that does not support the requirements needed for the extra features.

> [!TIP]
> Do not repeat **Requirement**s from the base manifest inside a **VersionOverrides**. Doing so has no effect and is potentially misleading as to the purpose of the **Requirements** element inside a **VersionOverrides**.

> [!WARNING]
> You should use great care before using a **Requirements** element in a **VersionOverrides**, because on platform and version combinations that don't support the requirement, *none* of the add-in commands will be installed, *even those that invoke functionality that doesn't need the requirement*. Consider, for example, an add-in that has two custom ribbon buttons. One of them calls Office JavaScript APIs that are available in requirement set **ExcelApi 1.4** (and later). The other calls APIs that are only available in **ExcelApi 1.9** (and later). If you put a requirement for **ExcelApi 1.9** in the **VersionOverrides**, *neither* button will appear on the ribbon. A better strategy in this scenario would be to use the technique described in [Runtime checks for method and requirement set support](../../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). The code invoked by the second button first uses `isSetSupported` to check for support of **ExcelApi 1.9**. If it is not supported, the code gives the user a message saying that this feature of the add-in is not available on their version of Office. 

> [!NOTE]
> In Mail add-ins, it is possible for a **VersionOverrides** 1.1 to be nested inside a **VersionOverrides** 1.0. Office will always use the highest version **VersionOverrides** that is supported by the platform and Office version.

## Syntax

```XML
<Requirements>
   ...
</Requirements>
```

## Contained in

[OfficeApp](officeapp.md)
[VersionOverrides](versionoverrides.md)

## Can contain

|Element|Content|Mail|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Methods](methods.md)|x||x|

## See also

For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).
