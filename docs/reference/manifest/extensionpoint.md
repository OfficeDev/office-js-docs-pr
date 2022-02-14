---
title: ExtensionPoint element in the manifest file
description: Defines where an add-in exposes functionality in the Office UI.
ms.date: 02/11/2022
ms.localizationpriority: medium
---

# ExtensionPoint element

 Defines where an add-in exposes functionality in the Office UI. The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Yes  | The type of extension point being defined. Possible values depend on the Office host application defined in the grandparent **Host** element value.|

## Extension points for Excel, OneNote, PowerPoint, and Word add-in commands

There are three types of extension points available in some or all of these hosts.

- [PrimaryCommandSurface](#primarycommandsurface) (Valid for Word, Excel, PowerPoint, and OneNote) - The ribbon in Office.
- [ContextMenu](#contextmenu) (Valid for Word, Excel, PowerPoint, and OneNote) - The shortcut menu that appears when you select and hold (or right-click) in the Office UI.
- [CustomFunctions](#customfunctions) (Valid only for Excel) - A custom function written in JavaScript for Excel.

See the following subsections for the child elements and examples of these types of extension points.

### PrimaryCommandSurface

The primary command surface in Word, Excel, PowerPoint, and OneNote is the ribbon.

#### Child elements

|Element|Description|
|:-----|:-----|
|[CustomTab](customtab.md|Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|[OfficeTab](officetab.md)|Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**). If you use the **OfficeTab** element, you can't use the **CustomTab** element.|

#### Example

The following example shows how to use the **ExtensionPoint** element with **PrimaryCommandSurface**. It adds a custom tab to the ribbon.

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID.

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### ContextMenu

A context menu is a shortcut menu that appears when you right-click in the Office UI.

#### Child elements
 
|Element|Description|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to one of the following strings: <br/> - **ContextMenuText** if the context menu should open when a user right-clicks on the selected text. <br/> - **ContextMenuCell** if the context menu should open when the user right-clicks on a cell on an Excel spreadsheet.|

#### Example

The following adds a custom context menu to the cells in an Excel spreadsheet.

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

### CustomFunctions

A custom function written in JavaScript or TypeScript for Excel.

#### Child elements

|Element|Description|
|:-----|:-----|
|[Script](script.md)|Required. Links to the JavaScript file with the custom function's definition and registration code.|
|[Page](page.md)|Required. Links to the HTML page for your custom functions.|
|[MetaData](metadata.md)|Required. Defines the metadata settings used by a custom function in Excel.|
|[Namespace](namespace.md)|Optional. Defines the namespace used by a custom function in Excel.|

#### Example

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

## Extension points for Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### MessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

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
  <CustomTab id="Contoso.TabCustom2">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### MessageComposeCommandSurface

This extension point puts buttons on the ribbon for add-ins using mail compose form. 

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

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
  <CustomTab id="Contoso.TabCustom3">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentOrganizerCommandSurface

This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting. 

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

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
  <CustomTab id="Contoso.TabCustom4">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting. 

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

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
  <CustomTab id="Contoso.TabCustom5">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

This extension point puts buttons on the ribbon for the module extension.

> [!IMPORTANT]
> Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

### MobileMessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view in the mobile form factor.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [Group](group.md) |  Adds a group of buttons to the command surface.  |

**ExtensionPoint** elements of this type can only have one child element: a **Group** element.

**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.

#### Example

```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
      <Control  xsi:type="MobileButton id="Contoso.mobileButton1"">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### MobileOnlineMeetingCommandSurface

This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor. A meeting organizer can create an online meeting. An attendee can subsequently join the online meeting. To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.

> [!NOTE]
> This extension point is only supported on Android and iOS with a Microsoft 365 subscription.
>
> Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [Control](control.md) |  Adds a button to the command surface.  |

`ExtensionPoint` elements of this type can only have one child element: a `Control` element.

The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.

The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).

#### Example

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### LaunchEvent

This extension point enables an add-in to activate based on supported events in the desktop form factor. To learn more about this scenario and for the full list of supported events, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.

> [!IMPORTANT]
> Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  List of [LaunchEvent](launchevent.md) for event-based activation.  |
| [SourceLocation](sourcelocation.md) |  The location of the source JavaScript file.  |

#### Example

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### Events

This extension point adds an event handler for a specified event. For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).

> [!IMPORTANT]
> Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

| Element | Description  |
|:-----|:-----|
|  [Event](event.md) |  Specifies the event and event handler function.  |

#### ItemSend event example

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### DetectedEntity

This extension point adds a contextual add-in activation on a specified entity type.

> [!IMPORTANT]
> Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.

> [!NOTE]
> This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

|  Element |  Description  |
|:-----|:-----|
|  [Label](#label) |  Specifies the label for the add-in in the contextual window.  |
|  [SourceLocation](sourcelocation.md) |  Specifies the URL for the contextual window.  |
|  [Rule](rule.md) |  Specifies the rule or rules that determine when an add-in activates.  |

#### Label

Required. The label of the group. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

#### Highlight requirements

The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.

However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.

- The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.
- If using a single rule, `Highlight` MUST be set to `all`.
- If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.
- If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.

#### DetectedEntity event example

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
