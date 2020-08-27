---
title: VersionOverrides element in the manifest file
description: 'Reference documentation of the VersionOverrides element for Office Add-ins manifest (XML) files.'
ms.date: 03/05/2020
localization_priority: Normal
---

# VersionOverrides element

The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Yes  |  The VersionOverrides schema namespace. The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element. See [Namespace values](#namespace-values) below.|
|  **xsi:type**  |  Yes  | The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`. |

### Namespace values

The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.

- **TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.
- **ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.
- **MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:
    - When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.
    - When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.

> [!NOTE]
> Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Description**    |  No   |  Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.|
|  **Requirements**  |  No   |  Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.|
|  [Hosts](hosts.md)                |  Yes  |  Specifies a collection of Office applications. The child Hosts element overrides the Hosts element in the parent portion of the manifest.  |
|  [Resources](resources.md)    |  Yes  | Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.|
|  [EquivalentAddins](equivalentaddins.md)    |  No  | Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in. The web add-in is not activated if an equivalent native add-in is installed.|
|  **VersionOverrides**    |  No  | Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  No  | Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  No  |  Specifies a collection of extended permissions.<br><br>**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment. |

### VersionOverrides example

The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
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

## Implementing multiple versions

A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.

In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version. The child `VersionOverrides` element doesn't inherit any values from the parent.

To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
