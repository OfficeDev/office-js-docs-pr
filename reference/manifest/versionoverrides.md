# VersionOverrides element

The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 schema. 

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xmlns](#xmlns)       |  Yes  |  The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides`.|
|  [xsi:type](#xsitype)  |  Yes  | The schema version. The version described in this article is `VersionOverridesV1_0`.|


## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Description](#description)    |  No   |  Describes the add-in.|
|  [Requirements](#requirements)  |  No   |  Specifies the minimum requirement set and version of Office.js that the add-in requires.| 
|  [Hosts](./hosts.md)                |  Yes  |  Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.  |
|  [Resources](./resources.md)    |  Yes  | Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.|


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
