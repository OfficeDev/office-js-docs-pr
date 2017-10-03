# VersionOverrides element

The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Yes  |  The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.|
|  **xsi:type**  |  Yes  | The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`. |

> **Note:** Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Description**    |  No   |  Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.|
|  **Requirements**  |  No   |  Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.|
|  [Hosts](./hosts.md)                |  Yes  |  Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.  |
|  [Resources](./resources.md)    |  Yes  | Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.|
|  **VersionOverrides**    |  No  | Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing_multiple_versions) for details. |
|  **WebApplicationInfo**    |  No  | Specifies details about the add-in's associated Web application. |



### VersionOverrides example
```xml
<OfficeApp>
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
<OfficeApp>
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
...
</OfficeApp>
```
