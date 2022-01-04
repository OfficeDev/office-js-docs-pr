---
title: VersionOverrides element in the manifest file
description: 'Reference documentation of the VersionOverrides element for Office Add-ins manifest (XML) files.'
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# VersionOverrides element

This element contains information for features that aren't supported in the base manifest. Its child markup may override some of the markup in the base manifest (or in a parent **VersionOverrides**). **VersionOverrides** is a child element of either the root [OfficeApp](officeapp.md) element in the manifest or a parent **VersionOverrides** element. This element is supported in manifest schema v1.1 and later but is defined in separate VersionOverrides schemas.

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Yes  |  The VersionOverrides schema namespace. The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element. See [Namespace values](#namespace-values) below.|
|  **xsi:type**  |  Yes  | The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`. |

### Namespace values

The following lists the required value of the **xmlns** attribute depending on the **xsi:type** value of the root `<OfficeApp>` element.

- **TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** must be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.
- **ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** must be `http://schemas.microsoft.com/office/contentappversionoverrides`.
- **MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:
  - When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.
  - When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.

> [!NOTE]
> Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.

## Variant schemas

There is a different schema for each of the possible **xmlns** values, so each has a separate reference page.

- [VersionOverrides 1.0 TaskPane](versionoverrides-1-0-taskpane.md)
- [VersionOverrides 1.0 Content](versionoverrides-1-0-content.md)
- [VersionOverrides 1.0 Mail](versionoverrides-1-0-mail.md)
- [VersionOverrides 1.1 Mail](versionoverrides-1-1-mail.md)
