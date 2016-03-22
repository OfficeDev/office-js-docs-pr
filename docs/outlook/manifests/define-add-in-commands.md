
# Define add-in commands in your Outlook add-in manifest

To support add-in commands, some additional elements have been added to the add-in manifest v1.1 within the  **VersionOverrides** element. When a manifest contains the **VersionOverrides** element, versions of Outlook that support add-in commands will use the information within that element to load the add-in. Earlier versions of Outlook that do not support add-in commands will ignore the element and continue to use the elements as described in [Outlook add-in manifests](../../outlook/manifests/manifests.md).

When the client application recognizes the  **VersionOverrides** node, the add-in name appears in the ribbon, not in the read/compose pane. The add-in won't appear in both places.
 

## VersionOverrides element

The  **VersionOverrides** element is the root element that contains information for the add-in commands implemented by the add-in. It is supported in manifest schema v1.1 or later but is defined in the VersionOverrides v1.0 schema. The attributes for **VersionOverrides** are as follows.

|**Attribute**|**Description**|
|:-----|:-----|
|**xmlns**| Required. The schema location. Must be "http://schemas.microsoft.com/office/mailappversionoverrides".|
|**xsi:type**|Required. The schema version. The version described in this topic is "VersionOverridesV1_0".|
The following table shows the child elements of  **VersionOverrides**.


|**Element**|**Description**|
|:-----|:-----|
|**Description**|Describes the add-in. This overrides the  **Description** element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the **Resources** element. The **resid** attribute of the **Description** element is set to the value of the **id** attribute of the **String** element that contains the text.|
|**Requirements**|Specifies the minimum requirement set and version of Office.js that the Office add-in needs to activate. It is defined the same as in [Outlook add-in manifests](../../outlook/manifests/manifests.md). This overrides the  **Requirements** element in the parent portion of the manifest.|
|**Hosts**|Required. Specifies a collection of host types and their settings. It overrides the  **Hosts** element in the parent portion of the manifest. It must have an **xsi:type** attribute set to "MailHost", and it must contain a **FormFactor** child element.|
|**Resources**|Defines a collection of resources (strings, URLs, and images) that are referenced by other elements of the manifest. This is described in the [Resources element](#VersionOverrides10_Resources) section later in this topic.|

Here an example of  **VersionOverrides**, showing its child elements.

```XML
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

## FormFactor element

The  **FormFactor** element specifies the settings for an add-in for a given form factor. It is a child node under **Hosts** / **Host**. Currently, it can only specify the desktop ( **DesktopFormFactor**). It contains all the add-in information for that form factor except for the  **Resources** node.

The form factor contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information see the following [FunctionFile element](#VersionOverrides10_FunctionFile) and [ExtensionPoint element](#VersionOverrides10_ExtensionPoint) sections. The following is an example of **FormFactor**, showing its child nodes.

```XML
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
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
  </VersionOverrides>
...
</OfficeApp>
```


## FunctionFile element


The  **FunctionFile** element is a child element under **FormFactor**. It specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all of the JavaScript functions used by UI-less add-in command buttons. For more information, see the [Button controls](#VersionOverrides10_Buttons) section of this article.

The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`. The functions should use the [item.notificationMessages](../../../reference/outlook/Office.context.mailbox.item.md) API to indicate progress, success, or failure to the user. It should also call [event.completed (JavaScript API for Office)](../../../reference/shared/event.completed.md) when it has finished execution. The name of the functions are used in the **FunctionName** element for UI-less buttons.

The following is an example of an HTML file defining a trackMessage function.

```js
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


## ExtensionPoint element


The  **ExtensionPoint** element defines where an add-in exposes functionality. It is a child element under **FormFactor**. For each form factor, you can define **ExtensionPoint** elements with the following **xsi:type** values:


- CustomPane 
    
- MessageReadCommandSurface 
    
- MessageComposeCommandSurface 
    
- AppointmentOrganizerCommandSurface 
    
- AppointmentAttendeeCommandSurface
    

### CustomPane

The  **CustomPane** extension point defines an add-in that activates when specified rules are satisfied. It is only for read form and it displays in a horizontal pane. The following are the elements of the **CustomPane**.

|**Element**|**Description**|
|:-----|:-----|
|**RequestedHeight**| Optional. The requested height, in pixels, for the display pane when it is running on a desktop computer. This can be from 32 to 450 pixels. It is the same as in read add-ins (see[RequestedHeight element (ItemReadTabletMailAppSettings complexType) (app manifest schema v1.1)](http://msdn.microsoft.com/library/6296f5b0-3d5b-5ab9-eee9-55a7eb90f92c%28Office.15%29.aspx)|
|**SourceLocation**|Required. The URL for the source code file of the add-in. This refers to a  **Url** element in the **Resources** element.|
|**Rule**|Required. The rule or collection of rules that specify when the add-in activates. It is the same as defined in [Outlook add-in manifests](../../outlook/manifests/manifests.md), except the [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) rule has the following changes: **ItemType** is either "Message" or "AppointmentAttendee", and there is no **FormType** attribute. For more information, see [Custom pane Outlook add-ins](../../outlook/custom-pane-outlook-add-ins.md) and [Activation rules for Outlook add-ins](../../outlook/manifests/activation-rules.md).|
|**DisableEntityHighlighting**|Optional. Specifies whether entity highlighting should be turned off for this mail add-in. |

The following example defines a custom pane for items that are messages or have an attachment or include an address.


```XML
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

On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in. If adding to the default tab, this is limited to one group per add-in. On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.

An example of a group on the default ribbon tab is as follows.

```XML
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

Where:

|**Element**|**Description**|
|:-----|:-----|
|**OfficeTab**|Required. The pre-existing tab to use. Currently, the  **id** attribute can only be "TabDefault".|
|**Group**|A group of user interface extension points in a tab. A group can have up to six controls.The  **id** attribute is required. It is a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|
|**Control**|A group requires at least one control. Currently, only buttons and menus are supported. See the following [Button controls](#VersionOverrides10_Buttons) and [Menu (dropdown button) controls](#VersionOverrides10_Menus) sections for more information.|

You can also create a custom tab on the ribbon by using the  **CustomTab** element, as shown in the following example.

```XML
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

Where:

|**Element**|**Description**|
|:-----|:-----|
|**CustomTab**|Required. The  **id** attribute must be unique within the manifest.|
|**Group**|A group of user interface extension points in a tab. A group can have up to six controls.The  **id** attribute is required. It is a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|
|**Control**|A group requires at least one control. Currently, only buttons and menus are supported. See the following [Button controls](#VersionOverrides10_Buttons) and[Menu (dropdown button) controls](#VersionOverrides10_Menus) sections for more information.|
|**Label**|Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|

#### Button controls


A button performs a single action when the user selects it. It can either execute a function or show a task pane.

The button control looks like the following:

```XML
<Control xsi:type="Button" id="<choose a descriptive name>" >
  <!-- include button elements, as described in the following table -->
</Control>
```

Where the  **id** attribute is a string with a maximum of 125 characters and the button elements are described in the following table.

|**Button elements**|**Description**|
|:-----|:-----|
|**Label**|Required. The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|
|**Supertip**|Required. The supertip for this button, which is defined by the following table.|

|**Element**|**Description**|
|:-----|:-----|
|**Title**|Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|
|**Description**|Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the **Resources** element.|
|**Icon**|Required. Contains the  **Image** elements for the button.|
|**Image**|An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|
|**Action**|Required. Specifies the action to perform when the user selects the button. It is defined by the following.<br>**xsi:type** This attribute specifies the kind of action performed when the user selects the button. It can be one of the following<ul><li><p>"ExecuteFunction"</p></li><li><p>"ShowTaskpane"</p></li></ul>|
|**FunctionName**|Required element when  **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the **FunctionFile** element.|
|**SourceLocation**|Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element.|

The following is an example of a  _UI-less button_, which executes a function named `getSubject`.

```XML
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

A  _task pane button_ control is a button that launches a task pane. Task pane buttons do not support toggles. The following is an example of a task pane button.




```XML
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


#### Menu (dropdown button) controls


A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported. 

The syntax for the menu control is as follows:




```XML
<Control xsi:type="Menu" id="<choose a descriptive name>" >
  <!-- include menu elements, as described in the following table -->
</Control>
```

Where the  **id** attribute is a string with a maximum of 125 characters and the menu elements are described in the following table.

|**Menu elements**|**Description**|
|:-----|:-----|
|**Label**|Required. The text for the menu. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|
|**SuperTip**|Required. The supertip for the menu, which is defined by the following table.|


|**Element**|**Description**|
|:-----|:-----|
|**Title**|Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the **Resources** element.|
|**Description**|Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the **Resources** element.|
|**Icon**|Required. Contains the  **Image** elements for the menu.|
|**Image**|An image for the menu. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|
|**Items**|Required. Contains the  **Item** elements for the menu. Each **Item** element contains the same child elements as a [Button controls](#VersionOverrides10_Buttons).|


The following is an example of a menu.

```XML
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


### MessageComposeCommandSurface

This puts buttons on the ribbon for add-ins using mail compose form. It is defined the same as for MessageReadCommandSurface.


### AppointmentOrganizerCommandSurface

This puts buttons on the ribbon for the form that's displayed to the organizer of the meeting. It is defined the same as for MessageReadCommandSurface.


### AppointmentAttendeeCommandSurface

This puts buttons on the ribbon for the form that's displayed to the attendee of the meeting. It is defined the same as for MessageReadCommandSurface.


## Resources element


The  **Resources** element contains icons, strings, and URLs for the **VersionOverrides** node. A manifest element specifies a resource by using the **Id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **Id** has a maximum of 32 characters.

The  **Resources** node defines the following resources. Each resource can have one or more **Override** child elements to define a resource for specific locales.

|**Resource**|**Description**|
|:-----|:-----|
|**Images**/ **Image**|Provides the HTTPS URL to an image for an icon. Each icon must have three  **Image** elements, one for each of the three mandatory sizes:<br><ul><li><p>16x16</p></li><li><p>32x32</p></li><li><p>80x80</p></li></ul>The following additional sizes are also supported, but not required:<ul><li><p>20x20</p></li><li><p>24x24</p></li><li><p>40x40</p></li><li><p>48x48</p></li><li><p>64x64</p></li></ul>|
|**Urls**/ **Url**|Provides an HTTPS URL location. A URL can be a maximum of 2048 characters. |
|**ShortStrings**/ **String**|The text for  **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.|
|**LongStrings**/ **String**|The text for  **Description** attributes. Each **String** contains a maximum of 250 characters.|

**Note**  When defining resources, keep the following requirements in mind:

The following is an example of the  **Resources** element.


```XML
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


## Rule changes


The following changes affect the rules in the manifest.


- Activation rules are now inside each entry point.
    
- [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) is modified so that **ItemType** is either Message or AppointmentAttendee and there is no **FormType** attribute.
    
- [ItemHasKnownEntity](http://msdn.microsoft.com/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) Is modified to accept a string for entity type, rather than an enum.
    

## Sample Manifest


For a full sample manifest, see the [command-demo](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml%28Office.15%29.aspx) sample on GitHub.


## Additional resources



- [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md)
    
- [Outlook add-in manifests](../../outlook/manifests/manifests.md)
    
- [command-demo sample](https://github.com/jasonjoh/command-demo.aspx)
    
