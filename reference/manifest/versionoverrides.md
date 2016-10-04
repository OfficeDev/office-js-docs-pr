# VersionOverrides element

The root element that contains information for the add-in commands implemented by the add-in. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 schema. 

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xmlns](#xmlns)       |  Yes  |  The schema location. Must be `http://schemas.microsoft.com/office/mailappversionoverrides`.|
|  [xsi:type](#xsitype)  |  Yes  | The schema version. The version described in this topic is `VersionOverridesV1_0`.|


## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Description](#description)    |  No   |  Describes the add-in. |
|  [Requirements](#requirements)  |  No   |  The minimum Mailbox version required. | 
|  [Hosts](./hosts.md)                |  Yes  |  A collection of host types and their settings. |
|  [Resources](./resources.md)    |  Yes  | Resource definitions (strings, URLs, and images).  |


### xmlns 
This is a required attribute that defines the schema location. The value should always be defined as `http://schemas.microsoft.com/office/mailappversionoverrides`.

### xsi:type
This is a required attribute which defines the schema version. At this time the only valid value is `VersionOverridesV1_0`.  

### Description
Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.

### Requirements
Specifies the minimum requirement set and version of Office.js that the Office Add-in needs to activate. This overrides the  `Requirements` element in the parent portion of the manifest.

### Hosts
See [Hosts](./hosts.md).

### Resources 
See [Resources](./resources.md).


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
