# Hosts element

Specifies the Office client application where the Office add-in will activate. Contains a collection of host objects and their settings. 

When included in the [VersionOverrides](./versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest. 

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Host](#host-element)    |  Yes   |  Describes a host and its settings. |

> ** Note: ** Outlook requires `Hosts`to contain a `Host` definition for `MailHost`.

---- 

## Host element
Specifies an individual Office application type where the add-in should activate, such as “document”, “workbook”, “presentation”, “project”, “mailbox”.

### Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Yes  | Describes the Office host these settings apply to.|

### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  Yes   |  Defines the form factor affected. |


### xsi:type
Controls which Office host (Word, Excel, PowerPoint, Outlook) the contained settings apply too. The value must be one of the following:

- `MailHost` (Outlook)    


### FormFactor
See [FormFactor](./formfactor.md).


## Hosts Example 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
