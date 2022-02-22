---
title: VersionOverrides 1.0 element in the manifest file for a content add-in
description: Reference documentation of the VersionOverrides element (content) for Office Add-ins manifest (XML) files.
ms.date: 02/18/2022
ms.localizationpriority: medium
---

# VersionOverrides 1.0 element in the manifest file for a content add-in

This element contains information for features that aren't supported in the base manifest.

> [!NOTE]
> This article assumes that you're familiar with the [overview of the VersionOverrides element](versionoverrides.md), which contains important information about the element's attributes and variations.

## Child elements

The following table applies only to version 1.0 of **VersionOverrides** elements and only to content add-ins.

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  No  | Not currently usable in VersionOverrides 1.0 for content add-ins. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  No  | Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0. |

## Example

The following is a simple example. For more complex examples, see the manifests for the sample add-ins in [Office Add-in code samples](https://github.com/OfficeDev/PnP-OfficeAddins).

```xml
<OfficeApp ... xsi:type="Content">
...
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/contentappversionoverrides" xsi:type="VersionOverridesV1_0">
        <WebApplicationInfo>
            <Id>$application_GUID here$</Id>
            <Resource>api://localhost:44355/$application_GUID here$</Resource>
            <Scopes>
                <Scope>Files.Read.All</Scope>
                <Scope>profile</Scope>
            </Scopes>
        </WebApplicationInfo>
    </VersionOverrides>
...
</OfficeApp>
```
