# ExtensionPoint element

 Defines where an add-in exposes functionality in the Office UI. The **ExtensionPoint** element is a child element of [FormFactor](./formfactor.md). 

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Yes  | The type of extension point being defined.|


## Extension points for Word, Excel, PowerPoint, and OneNote add-in commands

- **PrimaryCommandSurface** - The ribbon in Office.
- **ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.

The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.


 >**Important**  For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format.<CustomTab id="mycompanyname.mygroupname">


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

**Child elements**
 
|**Element**|**Description**|
|:-----|:-----|
|**CustomTab**|Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|**OfficeTab**|Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).|
|**OfficeMenu**|Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: <br/> - **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> - **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.|
|**Group**|A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.|
|**Icon**|Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.|
|**Tooltip**|Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.|
|**Control**|Each group requires at least one control. A  **Control** element can be either a **Button** or a **Menu**. Use  **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported.See the [Button controls](#button-controls) and [Menu controls](#menu-controls) sections for more information.<br/>**Note**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.

## Extension points for Outlook add-in commands

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Can only be used in the [DesktopFormFactor](./formfactor.md).)

### CustomPane

The CustomPane extension point defines an add-in that activates when specified rules are satisfied. It is only for read form and it displays in a horizontal pane. 

**Child elements**

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **RequestedHeight** | No |  The requested height, in pixels, for the display pane when it is running on a desktop computer. This can be from 32 to 450 pixels.  |
|  **SourceLocation**  | Yes |  The URL for the source code file of the add-in. This refers to a  **Url** element in the [Resources](./resources.md)  element.  |
|  **Rule**  | Yes |  The rule or collection of rules that specify when the add-in activates. For more information, see  [Activation rules for Outlook add-ins](../../docs/outlook/manifests/activation-rules.md). |
|  **DisableEntityHighlighting**  | No |  Specifies whether entity highlighting should be turned off. |


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

**Child elements**

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

**Child elements**

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

**Child elements**

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

**Child elements**

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

**Child elements**

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

