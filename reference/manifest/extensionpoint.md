# ExtensionPoint element

 Defines where an add-in exposes functionality. The **ExtensionPoint** element is a child element of [FormFactor](./formfactor.md). 

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Yes  | The type of ExtentionPoint being defined.|

## xsi:type
For each form factor, you can define **ExtensionPoint** elements with the following **xsi:type** values, with the exception of the **Module** value which can only be used in the [DesktopFormFactor](./formfactor.md):

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module)

### CustomPane

The CustomPane extension point defines an add-in that activates when specified rules are satisfied. It is only for read form and it displays in a horizontal pane. 

#### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [RequestedHeight](#requestedheight) | No |  The requested height in pixels.  |
|  [SourceLocation](#sourcelocation)  | Yes |  The URL for the source code file of the add-in.  |
|  [Rule](#rule)  | Yes |  The rule or collection of rules that specify when the add-in activates.  |
|  [DisableEntityHighlighting](#disableentityhighlighting)  | No |  Specifies whether entity highlighting should be turned off. |

#### RequestedHeight
Optional. The requested height, in pixels, for the display pane when it is running on a desktop computer. This can be from 32 to 450 pixels. It is the same as in read add-ins (see [RequestedHeight element](../reference/requestedheight.md)

#### SourceLocation
Required. The URL for the source code file of the add-in. This refers to a  **Url** element in the [Resources](./resources.md)  element.

#### Rule
Required. The rule or collection of rules that specify when the add-in activates. It is the same as defined in [Outlook add-in manifests](../../outlook/manifests/manifests.md), except the ItemIs rule has the following changes: **ItemType** is either "Message" or "AppointmentAttendee", and there is no **FormType** attribute. For more information, see [Custom pane Outlook add-ins](../../outlook/custom-pane-outlook-add-ins.md) and [Activation rules for Outlook add-ins](../../outlook/manifests/activation-rules.md).

#### DisableEntityHighlighting
Optional. Specifies whether entity highlighting should be turned off for this Outlook add-in. 

#### CustomPane example
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

#### Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
This extension point puts buttons on the ribbon for add-ins using mail compose form. 

#### Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting. 

#### Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting. 

#### Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

This extension point puts buttons on the ribbon for the module extension. 

#### Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

