---
title: Create add-in commands in your manifest for Excel, PowerPoint, and Word
description: Use VersionOverrides in your manifest to define add-in commands for Excel, PowerPoint, and Word. Use add-in commands to create UI elements, add buttons or lists, and perform actions.
ms.date: 07/18/2022
ms.localizationpriority: medium
---


# Create add-in commands in your manifest for Excel, PowerPoint, and Word

> [!NOTE]
> Add-in commands are also supported in Outlook. For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md)

Use **[VersionOverrides](/javascript/api/manifest/versionoverrides)** in your manifest to define add-in commands for Excel, PowerPoint, and Word. Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. For an introduction to add-in commands, see [Add-in commands for Excel, PowerPoint, and Word](../design/add-in-commands.md).

This article describes how to edit your manifest to define add-in commands and how to create the code for [function commands](../design/add-in-commands.md#types-of-add-in-commands). The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.

![Overview of add-in commands elements in the manifest. The top node here is VersionOverrides with children Hosts and Resources. Under Hosts are Host then DesktopFormFactor. Under DesktopFormFactor are FunctionFile and ExtensionPoint. Under ExtensionPoint are CustomTab or OfficeTab and Office Menu. Under CustomTab or Office Tab are Group then Control then Action. Under Office Menu are Control then Action. Under Resources (child of VersionOverrides) are Images, Urls, ShortStrings, and LongStrings.](../images/version-overrides.png)

## Step 1: Create the project

We recommend you create a project by following one of the quick starts such as [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md). Each quick start for Excel, Word, and PowerPoint generates a project that already contains an add-in command (button) to show the task pane. Ensure that you have read [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.

## Step 2: Create a task pane add-in

To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article. You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropriate **XML namespaces** as well as add the **\<VersionOverrides\>** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).

The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **\<VersionOverrides\>** element. Office 2013 doesn't support add-in commands, but by adding **\<VersionOverrides\>** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **\<SourceLocation\>** to run your add-in as a single task pane add-in. In Office 2016, if no **\<VersionOverrides\>** element is included, the task pane of your add-in automatically opens to the URL specified in **\<SourceLocation\>**. If you include **\<VersionOverrides\>**, however, your add-in displays the add-in commands only, and doesn't initially display your add-in's task pane.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="https://www.contoso.com/Images/Icon_32.png" />
  <SupportUrl DefaultValue="https://www.contoso.com/contact" />
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## Step 3: Add VersionOverrides element

The **\<VersionOverrides\>** element is the root element that contains the definition of your add-in command. **\<VersionOverrides\>** is a child element of the **\<OfficeApp\>** element in the manifest. The following table lists the attributes of the **\<VersionOverrides\>** element.

|Attribute|Description|
|:-----|:-----|
|**xmlns** <br/> | Required. The schema location, which must be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`. <br/> |
|**xsi:type** <br/> |Required. The schema version. The version described in this article is "VersionOverridesV1_0".  <br/> |

The following table identifies the child elements of **\<VersionOverrides\>**.
  
|Element|Description|
|:-----|:-----|
|**\<Description\>** <br/> |Optional. Describes the add-in. This child **\<Description\>** element overrides a previous **\<Description\>** element in the parent portion of the manifest. The **resid** attribute for this **\<Description\>** element is set to the **id** of a **\<String\>** element. The **\<String\>** element contains the text for **\<Description\>**. <br/> |
|**\<Requirements\>** <br/> |Optional. Specifies the minimum requirement set and version of Office.js that the add-in requires. This child **\<Requirements\>** element overrides the **\<Requirements\>** element in the parent portion of the manifest. For more information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**\<Hosts\>** <br/> |Required. Specifies a collection of Office applications. The child **\<Hosts\>** element overrides the **\<Hosts\>** element in the parent portion of the manifest. You must include a **xsi:type** attribute set to "Workbook" or "Document". <br/> |
|**\<Resources\>** <br/> |Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **\<Description\>** element's value refers to a child element in **\<Resources\>**. The **\<Resources\>** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. <br/> |

The following example shows how to use the **\<VersionOverrides\>** element and its child elements.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## Step 4: Add Hosts, Host, and DesktopFormFactor elements

The **\<Hosts\>** element contains one or more **\<Host\>** elements. A **\<Host\>** element specifies a particular Office application. The **\<Host\>** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office application. To show the same add-in commands in two or more different Office applications, you must duplicate the child elements in each **\<Host\>**.

The **\<DesktopFormFactor\>** element specifies the settings for an add-in that runs in Office on the web (in a browser) and Windows.

The following is an example of **\<Hosts\>**, **\<Host\>**, and **\<DesktopFormFactor\>** elements.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## Step 5: Add the FunctionFile element

The **\<FunctionFile\>** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](/javascript/api/manifest/control-button) for a description). The **\<FunctionFile\>** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **\<Url\>** element in the **\<Resources\>** element.

The following is an example of the **\<FunctionFile\>** element.
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint>

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> Make sure your JavaScript code calls  `Office.initialize`.

The JavaScript in the HTML file referenced by the **\<FunctionFile\>** element must call `Office.initialize`. The **\<FunctionName\>** element (see [Button controls](/javascript/api/manifest/control-button) for a description) uses the functions in **\<FunctionFile\>**.

The following code shows how to implement the function used by **\<FunctionName\>**.

```js
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Define the function.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("Function command works. Button ID=" + event.source.id,
            function (asyncResult) {
                const error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // Show error message.
                }
                else {
                    // Show success message.
                }
            });

        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
    }
    
    // You must register the function with the following line.
    Office.actions.associate("writeText", writeText);
</script>
```

> [!IMPORTANT]
> The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.

## Step 6: Add ExtensionPoint elements

The **\<ExtensionPoint\>** element defines where add-in commands should appear in the Office UI. You can define **\<ExtensionPoint\>** elements with these **xsi:type** values.

- **PrimaryCommandSurface**, which refers to the ribbon in Office.

- **ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.

The following examples show how to use the **\<ExtensionPoint\>** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.
  
```xml
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

|Element|Description|
|:-----|:-----|
|**\<CustomTab\>** <br/> |Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **\<CustomTab\>** element, you can't use the **\<OfficeTab\>** element. The **id** attribute is required. <br/> |
|**\<OfficeTab\>** <br/> |Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**). If you use the **\<OfficeTab\>** element, you can't use the **\<CustomTab\>** element. <br/> For more tab values to use with the **id** attribute, see [Tab values for default Office app ribbon tabs](/javascript/api/manifest/officetab).  <br/> |
|**\<OfficeMenu\>** <br/> | Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: <br/> **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> **ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. <br/> |
|**\<Group\>** <br/> |A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. <br/> |
|**\<Label\>** <br/> |Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<ShortStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Icon\>** <br/> |Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **\<Image\>** element. The **\<Image\>** element is a child element of the **\<Images\>** element, which is a child element of the **\<Resources\>** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. <br/> |
|**\<Tooltip\>** <br/> |Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<LongStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Control\>** <br/> |Each group requires at least one control. A **\<Control\>** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See [Button controls](/javascript/api/manifest/control-button) and [Menu controls](/javascript/api/manifest/control-menu) for more information. <br/>**Note:** To make troubleshooting easier, we recommend that you add a **\<Control\>** element and the related **\<Resources\>** child elements one at a time.          |

### Button controls

A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **\<Control\>** element:

- The **type** attribute is required, and must be set to **Button**.

- The **id** attribute of the **\<Control\>** element is a string with a maximum of 125 characters.

```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|Elements|Description|
|:-----|:-----|
|**\<Label\>** <br/> |Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<ShortStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Tooltip\>** <br/> |Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<LongStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Supertip\>** <br/> | Required. The supertip for this button, which is defined by the following: <br/> **Title** <br/>  Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<ShortStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> **\<Description\>** <br/>  Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<LongStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Icon\>** <br/> | Required. Contains the **\<Image\>** elements for the button. Image files must be .png format. <br/> **\<Image\>** <br/>  Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **\<Image\>** element. The **\<Image\>** element is a child element of the **\<Images\>** element, which is a child element of the **\<Resources\>** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. <br/> |
|**\<Action\>** <br/> | Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: <br/> **ExecuteFunction**, which runs a JavaScript function located in the file referenced by **\<FunctionFile\>**. The **\<FunctionName\>** child element specifies the name of the function to execute. <br/> **ShowTaskPane**, which shows the add-in's task pane. The **\<SourceLocation\>** child element specifies the source file location of the page to display. The **resid** attribute must be set to the value of the **id** attribute of a **\<Url\>** element in the **\<Urls\>** element in the **\<Resources\>** element. <br/> |

### Menu controls

A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:
  
- A root-level menu item.
- A list of submenu items.

When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.

The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **\<Control\>** element:

- The **xsi:type** attribute is required, and must be set to **Menu**.
- The **id** attribute is a string with a maximum of 125 characters.

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|Elements|Description|
|:-----|:-----|
|**\<Label\>** <br/> |Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<ShortStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Tooltip\>** <br/> |Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<LongStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<SuperTip\>** <br/> | Required. The supertip for the menu, which is defined by the following: <br/> **\<Title\>** <br/>  Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<ShortStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> **\<Description\>** <br/>  Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **\<String\>** element. The **\<String\>** element is a child element of the **\<LongStrings\>** element, which is a child element of the **\<Resources\>** element. <br/> |
|**\<Icon\>** <br/> | Required. Contains the **\<Image\>** elements for the menu. Image files must be .png format. <br/> **\<Image\>** <br/>  An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **\<Image\>** element. The **\<Image\>** element is a child element of the **\<Images\>** element, which is a child element of the **\<Resources\>** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. <br/> |
|**\<Items\>** <br/> |Required. Contains the **\<Item\>** elements for each submenu item. Each **\<Item\>** element contains the same child elements as [Button controls](/javascript/api/manifest/control-button).  <br/> |

## Step 7: Add the Resources element

The **\<Resources\>** element contains resources used by the different child elements of the **\<VersionOverrides\>** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.
  
The following shows an example of how to use the **\<Resources\>** element. Each resource can have one or more **\<Override\>** child elements to define a different resource for a specific locale.

```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|Resource|Description|
|:-----|:-----|
|**\<Images\>**/ **\<Image\>** <br/> | Provides the HTTPS URL to an image file. Each image must define the three required image sizes: <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  The following image sizes are also supported, but not required: <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**\<Urls\>**/ **\<Url\>** <br/> |Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.  <br/> |
|**\<ShortStrings\>**/ **\<String\>** <br/> |The text for **\<Label\>** and **\<Title\>** elements. Each **\<String\>** contains a maximum of 125 characters. <br/> |
|**\<LongStrings\>**/ **\<String\>** <br/> |The text for **\<Tooltip\>** and **\<Description\>** elements. Each **\<String\>** contains a maximum of 250 characters. <br/> |

> [!NOTE]
> You must use Secure Sockets Layer (SSL) for all URLs in the **\<Image\>** and **\<Url\>** elements.

### Tab values for default Office app ribbon tabs

In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **\<OfficeTab\>** element. The tab values are case sensitive.

|Office client application|Tab values|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## See also

- [Add-in commands for Excel, PowerPoint, and Word](../design/add-in-commands.md)
- [Sample: Create an Excel add-in with command buttons](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [Sample: Create an Word add-in with command buttons](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [Sample: Create an PowerPoint add-in with command buttons](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
