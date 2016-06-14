# Define add-in commands in your Outlook add-in manifest

To support add-in commands, some additional elements have been added to the add-in manifest v1.1 within the [VersionOverrides](../../../reference/manifest/commands/versionoverrides.md) element. When a manifest contains the **VersionOverrides** element, versions of Outlook that support add-in commands will use the information within that element to load the add-in. Earlier versions of Outlook that do not support add-in commands will ignore the element and continue to use the elements as described in [Outlook add-in manifests](../../outlook/manifests/manifests.md).

When the client application recognizes the  **VersionOverrides** node, the add-in name appears in the ribbon, not in the read/compose pane. The add-in won't appear in both places.
 

## VersionOverrides element

The  **VersionOverrides** element is the root element that contains information for the add-in commands implemented by the add-in. It is supported in manifest schema v1.1 or later but is defined in the VersionOverrides v1.0 schema. The attributes for **VersionOverrides** are as follows.

### Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xmlns](#xmlns)       |  Yes  |  The schema location. Must be `http://schemas.microsoft.com/office/mailappversionoverrides`|
|  [xsi:type](#xsitype)  |  Yes  | The schema version. The version described in this topic is `VersionOverridesV1_0`.|


### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Description](#description)    |  No   |  Describes the add-in. |
|  [Requirements](#requirements)  |  No   |  Minimum Mailbox version required | 
|  [Hosts](#hosts)                |  Yes  |  Collection of host types and their settings |
|  [Resources](#resources)|  Yes  | Resource definitions (strings, URLs, and images)  |


#### xmlns 
This is a required attribute which defines the schema location. The value should always be defined as `http://schemas.microsoft.com/office/mailappversionoverrides`.

#### xsi:type
This is a required attribute which defines the schema version. At this time the only valid value is `VersionOverridesV1_0`.  

#### Description
Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the LongString element contained in the [Resources](#eesources-element) element. The `resid` attribute of the Description element is set to the value of the `id` attribute of the `String` element that contains the text.

#### Requirements
Specifies the minimum requirement set and version of Office.js that the Office add-in needs to activate. It is defined the same as in [Outlook add-in manifests](../../outlook/manifests/manifests.md). This overrides the  `Requirements` element in the parent portion of the manifest.

#### Hosts
This contains a collection of host types and their settings. It overrides the  Hosts element in the parent portion of the manifest. It must have an [xsi:type](#xsitype) attribute set to "MailHost", and it must contain a `FormFactor` child element. This is a required element. 
```xml
<Hosts><Host xsi:type="MailHost"></Host></Hosts>
```

#### Resources 
This defines a collection of resources (strings, URLs, and images) that are referenced by other elements of the manifest. This is a required element. For more information, see [Resources Element](#eesources-element) 


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

---- 

## FormFactor element

The  **FormFactor** element specifies the settings for an add-in for a given form factor. It is a child node under **Hosts** / **Host**. Currently, it can only specify the desktop ( **DesktopFormFactor**). It contains all the add-in information for that form factor except for the  **Resources** node.

The form factor contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information see the following [FunctionFile element](#functionfile-element) and [ExtensionPoint element](#extensionpoint-element) sections. The following is an example of **FormFactor**, showing its child nodes.

### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [FunctionFile](#functionfile)      | Yes |  Url to file containing JavaScript functions  |
|  [ExtensionPoint](#extensionpoint)  | Yes |  Defines where an add-in exposes functionality  |

### FormFactor example
```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <ExtensionPoint xsi:type="CustomPane">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
---

## FunctionFile element

The  **FunctionFile** element is a child element under **FormFactor**. It specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all of the JavaScript functions used by UI-less add-in command buttons. For more information, see the [Button controls](#button-controls) section of this article.

The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`. The functions should use the [item.notificationMessages](../../../reference/outlook/Office.context.mailbox.item.md) API to indicate progress, success, or failure to the user. It should also call [event.completed](../../../reference/shared/event.completed.md) when it has finished execution. The name of the functions are used in the **FunctionName** element for UI-less buttons.

The following is an example of an HTML file defining a trackMessage function.

```javascript
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

---

## ExtensionPoint element

The  **ExtensionPoint** element defines where an add-in exposes functionality. It is a child element under **FormFactor**. For each form factor, you can define **ExtensionPoint** elements with the following **xsi:type** values, with the exception of the **Module** value which can only be used in the **DesktopFormFactor**:


- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module)
    
--- 

### CustomPane

The  **CustomPane** extension point defines an add-in that activates when specified rules are satisfied. It is only for read form and it displays in a horizontal pane. The following are the elements of the **CustomPane**.

### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [RequestedHeight](#requestedheight) | No |  The requested height in pixels.  |
|  [SourceLocation](#sourcelocation)  | Yes |  he URL for the source code file of the add-in.  |
|  [Rule](#rule)  | Yes |  The rule or collection of rules that specify when the add-in activates.  |
|  [DisableEntityHighlighting](#disableentityhighlighting)  | No |  Specifies whether entity highlighting should be turned off. |

#### RequestedHeight
Optional. The requested height, in pixels, for the display pane when it is running on a desktop computer. This can be from 32 to 450 pixels. It is the same as in read add-ins (see [RequestedHeight element](http://msdn.microsoft.com/library/6296f5b0-3d5b-5ab9-eee9-55a7eb90f92c%28Office.15%29.aspx)

#### SourceLocation
Required. The URL for the source code file of the add-in. This refers to a  **Url** element in the **Resources** element.

#### Rule
Required. The rule or collection of rules that specify when the add-in activates. It is the same as defined in [Outlook add-in manifests](../../outlook/manifests/manifests.md), except the [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) rule has the following changes: **ItemType** is either "Message" or "AppointmentAttendee", and there is no **FormType** attribute. For more information, see [Custom pane Outlook add-ins](../../outlook/custom-pane-outlook-add-ins.md) and [Activation rules for Outlook add-ins](../../outlook/manifests/activation-rules.md).

#### DisableEntityHighlighting
Optional. Specifies whether entity highlighting should be turned off for this mail add-in. 

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
----

### Office tab
On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in. 

The default tab is limited to one group per add-in. 

#### Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  OfficeTab  | Yes |  Always set to `TabDefault`.  |
|  Group      | Yes |  Defines a Group of commands.  |
|  Label      | Yes |  The label for the Group  |
|  Control    | Yes |  Collection of one or more Control objects  |

#### OfficeTab
Required. The pre-existing tab to use. Currently, the  **id** attribute can only be "TabDefault".

#### Group
A group of user interface extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manfiest. It is a string with a maximum of 125 characters.

#### Label
Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.

#### Control
A group requires at least one control. Currently, only buttons and menus are supported. See the following [Button controls](#button-controls) and [Menu (dropdown button) controls](#menu-dropdown-button-controls) sections for more information.

#### OfficeTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
      <Label resid="residTemplateManagement" />
      <Control xsi:type="Button" id="Button1" >
       <!-- information on the control -->
      </Control>
       <!-- other controls, as needed -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
----

### Custom tab
On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.

On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.

#### Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [CustomTab](#customtab)  | Yes |  Defines a custom ribbon tab.  |
|  [Group](#group)      | Yes |  Defines a Group of commands.  |
|  Label      | Yes |  The label for the CustomTab or a Group  |
|  [Control](#control)    | Yes |  Collection of one or more Control objects  |

#### CustomTab
Required. The  **id** attribute must be unique within the manifest.

#### Group
A group of user interface extension points in a tab. A group can have up to six controls.The  **id** attribute is required. It is a string with a maximum of 125 characters.

#### Label (Group)
Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.

#### Label (Tab)
Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.

#### Control
A group requires at least one control. Currently, only buttons and menus are supported. See the following [Button controls](#button-controls) and[Menu (dropdown button) controls](#menu-dropdown-button-controls) sections for more information.

####  CustomTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
      <Label resid="residCustomTabGroupLabel"/>
      <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
---- 

### MessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

---

### MessageComposeCommandSurface

This puts buttons on the ribbon for add-ins using mail compose form. It is defined the same as for MessageReadCommandSurface.

---

### AppointmentOrganizerCommandSurface

This puts buttons on the ribbon for the form that's displayed to the organizer of the meeting. It is defined the same as for MessageReadCommandSurface.

---

### AppointmentAttendeeCommandSurface

This puts buttons on the ribbon for the form that's displayed to the attendee of the meeting. It is defined the same as for MessageReadCommandSurface.

--- 

### Module

This puts buttons on the ribbon for the module extension. It is defined the same as for 
MessageReadCommandSurface.

---

### Button controls

A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest. 

#### Button control example
```xml
<Control xsi:type="Button" id="<choose a descriptive name>" >
  <!-- include button elements, as described in the following table -->
</Control>
```
#### Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  Label     | Yes |  The text for the button.         |
|  Supertip  | Yes |  The supertip for this button.    |
|  Icon      | Yes |  An image for the button.         |
|  Action    | Yes |  Specifies the action to perform  |

#### Label
Required. The text for the button. The  **resid** attribute must be set to the value of the **id** attribute 
of a **String** element in the **ShortStrings** element in the **Resources** element.

#### Supertip
Required. The Supertip object defines the tooltip (both Title and Description) attached to the Button. It has two required elements:

##### Title
Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.

##### Description
Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the **Resources** element.

#### Icon
Required. Contains the  **Image** elements for the button. 

##### Image
An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|

#### Action
Required. Specifies the action to perform when the user selects the button. It is defined by the following:

##### xsi:type
This attribute specifies the kind of action performed when the user selects the button. It can be one of the following
- ExecuteFunction
- ShowTaskpane

##### FunctionName
Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the **FunctionFile** element.

##### SourceLocation
Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element.

#### ExecuteFunction button example
```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>

```
#### ShowTaskpane button example
```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>

```

---

#### Menu (dropdown button) controls

A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported. 

The syntax for the menu control is as follows:




```XML
<Control xsi:type="Menu" id="<choose a descriptive name>" >
  <!-- include menu elements, as described in the following table -->
</Control>
```
#### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  Label     | Yes |  The text for the button.         |
|  Supertip  | Yes |  The supertip for this button.    |
|  Icon      | Yes |  An image for the button.         |
|  Action    | Yes |  Specifies the action to perform  |
|  Items     | Yes |  Collection of Buttons to display within the menu |

#### Label
Required. The text for the button. The  **resid** attribute must be set to the value of the **id** attribute 
of a **String** element in the **ShortStrings** element in the **Resources** element.

#### Supertip
Required. The Supertip object defines the tooltip (both Title and Description) attached to the Button. It has two required elements:

##### Title
Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.

##### Description
Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the **Resources** element.

#### Icon
Required. Contains the  **Image** elements for the button. 

##### Image
An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|

#### Action
Required. Specifies the action to perform when the user selects the button. It is defined by the following:

##### xsi:type
This attribute specifies the kind of action performed when the user selects the button. It can be one of the following
- ExecuteFunction
- ShowTaskpane

##### FunctionName
Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the **FunctionFile** element.

##### SourceLocation
Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element.

#### Items
Required. Contains the  **Item** elements for the menu. Each **Item** element contains the same child elements as a [Button controls](#button-controls).

#### Menu control example
```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

--- 

## Resources element

The **Resources** element contains icons, strings, and URLs for the [VersionOverrides](#versionoverrides) node. A manifest element specifies a resource by using the **Id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **Id** must be unique within the manifest and has a maximum of 32 characters.

The  **Resources** node defines the following resources. Each resource can have one or more **Override** child elements to define a resource for specific locales.

### Child elements

|  Element |  Type  |  Description  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  Provides the HTTPS URL to an image for an icon. |
|  [Urls](#urls)                |  url     |  Provides an HTTPS URL location. |
|  [ShotStrings](#shortstrings) |  string  |  The text for Label and Title elements. |
|  [LongStrings](#longstrings)  |  string  | The text for Description attributes. |

#### Images
Provides the HTTPS URL to an image for an icon. Each icon must have three  **Image** elements, one for each of the three mandatory sizes:
- 16x16
- 32x32
- 80x80

The following additional sizes are also supported, but not required:
- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> **Important: ** Outlook requires the ability to cache image resources for performance purposes. For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header. This will result in Outlook automatically subtituting a generic or default image.    

#### Urls
Provides an HTTPS URL location. A URL can be a maximum of 2048 characters. 

#### ShortStrings
The text for  **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.

#### LongStrings
The text for  **Description** attributes. Each **String** contains a maximum of 250 characters.

#### Resources example 
```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
---

## Rule changes

The following changes affect the rules in the manifest:

- Activation rules are now inside each entry point.
    
- [ItemIs](../../../reference/manifest/rule.md) is modified so that **ItemType** is either Message or AppointmentAttendee and there is no **FormType** attribute.
    
- [ItemHasKnownEntity](../../../reference/manifest/rule.md) Is modified to accept a string for entity type, rather than an enum.
    

## Sample manifest

For a full sample manifest, see the [Sample Outlook Manifest](https://gist.github.com/mlafleur/95b7ac030bb7a7ae742527e85a36b095) on GitHub.


## Additional resources



- [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md)
    
- [Outlook add-in manifests](../../outlook/manifests/manifests.md)
    
- [Outlook add-in command demo sample](https://github.com/jasonjoh/command-demo)
