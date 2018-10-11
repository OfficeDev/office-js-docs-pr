# ExtensionPoint element

 Defines where an add-in exposes functionality in the Office UI. The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). 

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Yes  | The type of extension point being defined.|

## Extension points for Excel only

- **CustomFunctions** - A custom function written in JavaScript for Excel.

[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.

## Extension points for Word, Excel, PowerPoint, and OneNote add-in commands

- **PrimaryCommandSurface** - The ribbon in Office.
- **ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.

The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.

> [!IMPORTANT] 
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### Child elements
 
|**Element**|**Description**|
|:-----|:-----|
|**CustomTab**|Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|**OfficeTab**|Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).|
|**OfficeMenu**|Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: <br/> - **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> - **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.|
|**Group**|A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.|
|**Icon**|Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.|
|**Tooltip**|Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.|
|**Control**|Each group requires at least one control. A  **Control** element can be either a **Button** or a **Menu**. Use  **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.<br/>**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.|
|**Script**|Links to the JavaScript file with the custom function definition and registration code. This element is not used in the Developer Preview. Instead, the HTML page is responsible for loading all JavaScript files.|
|**Page**|Links to the HTML page for your custom functions.|

## Extension points for Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
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
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### Events

This extension point adds an event handler for a specified event.

> [!NOTE]
> This element type is only supported by Outlook on the web in Office 365.

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

The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.

> [!NOTE]
> This element type is only supported by Outlook on the web in Office 365.

|  Element |  Description  |
|:-----|:-----|
|  [Label](#label) |  Specifies the label for the add-in in the contextual window.  |
|  [SourceLocation](sourcelocation.md) |  Specifies the URL for the contextual window.  |
|  [Rule](rule.md) |  Specifies the rule or rules that determine when an add-in activates.  |

#### Label

Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

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
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```