
# Host element
Specifies an individual Office application type where the add-in should activate.

> **Important**: There are two variation of the Host element depending on where it is defined; within the [basic manifest](#basic-manifest) or within [VersionOverrides](#versionoverrides). Both variations are covered here as they provide the same functionality but with slightly different syntax.  


## Basic Manifest

When defined in the basic manifest (under [OfficeApp](./officeapp.md)) the host type is determined by the `Name` attribute.   

### Attributes
| Attribute     | Type   | Required | Description                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | required | The name of the type of Office host application. |


### Name
Specifies the Host type targeted by this add-in. 

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### Example
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## VersionOverrides
When defined in [VersionOverrides](./versionoverrides) the host type is determined by the `xsi:type` attribute. 

### Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Yes  | Describes the Office host these settings apply to.|

### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  Yes   |  Defines the form factor affected. |


### xsi:type
Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) the contained settings apply too. The value must be one of the following:

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## Hosts example 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```