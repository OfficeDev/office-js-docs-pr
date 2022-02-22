---
title: VersionOverrides 1.0 element in the manifest file for a task pane add-in
description: Reference documentation of the VersionOverrides element (task pane) for Office Add-ins manifest (XML) files.
ms.date: 02/18/2022
ms.localizationpriority: medium
---

# VersionOverrides 1.0 element in the manifest file for a task pane add-in

This element contains information for features that aren't supported in the base manifest.

> [!NOTE]
> This article assumes that you're familiar with the [overview of the VersionOverrides element](versionoverrides.md), which contains important information about the element's attributes and variations.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (Required for Excel, PowerPoint, and Word.)
- Some child elements may be associated with additional requirement sets.

## Child elements

The following table applies only to version 1.0 of **VersionOverrides** elements and only to task pane add-ins.

> [!NOTE]
> In iOS, only **WebApplicationInfo** is supported. All other child elements of **VersionOverrides** are ignored.

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Description](#description)    |  No   |  Describes the add-in. |
|  [Requirements](requirements.md)  |  No   |  Specifies the minimum requirement sets that must be supported in order for the markup in the parent **VersionOverrides** to take effect. This should always be *more* restrictive than the **Requirements** element in the base portion of the manifest.|
|  [Hosts](hosts.md)                |  Yes  |  Specifies a collection of Office applications. The child Hosts element overrides the Hosts element in the parent portion of the manifest.  |
|  [Resources](resources.md)    |  Yes  | Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.|
|  [EquivalentAddins](equivalentaddins.md)    |  No  | Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in. The web add-in isn't activated if an equivalent native add-in is installed.|
|  **VersionOverrides**    |  No  | Not currently usable in VersionOverrides 1.0 for taskpane add-ins. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  No  | Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0. |

### Description

Describes the add-in. This overrides the **Description** element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element can be no more than 32 characters and must match the value of the `id` attribute of a child element of the **ShortString** element contained in the [Resources](resources.md) element.

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

## Example

The following is a simple example. For more complex examples, see the manifests for the sample add-ins in [Office Add-in code samples](https://github.com/OfficeDev/PnP-OfficeAddins).

```xml
<OfficeApp ... xsi:type="Taskpane">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```
