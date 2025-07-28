---
title: Create add-in commands with the add-in only manifest
description: Configure an add-in only manifest to define add-in commands for Excel, Outlook, PowerPoint, and Word. Use add-in commands to create UI elements, add buttons or menus, and perform actions.
ms.date: 02/28/2025
ms.localizationpriority: medium
---

# Create add-in commands with the add-in only manifest

Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. For an introduction to add-in commands, see [Add-in commands](../design/add-in-commands.md).

This article describes how to edit your add-in only manifest to define add-in commands and how to create the code for [function commands](../design/add-in-commands.md#types-of-add-in-commands).

> [!TIP]
> For instructions on how to create add-in commands with the unified manifest for Microsoft 365, see [Create add-in commands with the unified manifest for Microsoft 365](create-addin-commands-unified-manifest.md).

The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.

![Overview of add-in commands elements in the manifest. The top node here is VersionOverrides with children Hosts and Resources. Under Hosts are Host then DesktopFormFactor. Under DesktopFormFactor are FunctionFile and ExtensionPoint. Under ExtensionPoint are CustomTab or OfficeTab and Office Menu. Under CustomTab or Office Tab are Group then Control then Action. Under Office Menu are Control then Action. Under Resources (child of VersionOverrides) are Images, Urls, ShortStrings, and LongStrings.](../images/version-overrides.png)

## Sample commands

All the task pane add-ins created by [Yo Office](yeoman-generator-overview.md) have add-in commands. They contain an add-in command (button) to show the task pane. Generate these projects by following one of the quick starts, such as [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md). Ensure that you have read [Add-in commands](../design/add-in-commands.md) to understand command capabilities.

## Important parts of an add-in command

The following steps explain how to add add-in commands to an existing add-in.

### Step 1: Add VersionOverrides element

The [`<VersionOverrides>` element](/javascript/api/manifest/versionoverrides) is the root element that contains the definition of your add-in command. Details on the valid attributes and implications are found in [Version overrides in the manifest](xml-manifest-overview.md?tabs=tabid-1#version-overrides-in-the-manifest).

The following example shows the `<VersionOverrides>` element and its child elements.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <!-- Add information about requirement sets. -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- Add information about form factors. -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- Add information about resources. -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

### Step 2: Add Hosts, Host, and DesktopFormFactor elements

The [`<Hosts>` element](/javascript/api/manifest/hosts) contains one or more [`<Host>` elements](/javascript/api/manifest/host). A `<Host>` element specifies a particular Office application. The `<Host>` element contains child elements that specify the add-in commands to display after your add-in is installed in that Office application. To show the same add-in commands in two or more different Office applications, you must duplicate the child elements in each `<Host>`.

The [`<DesktopFormFactor>`](/javascript/api/manifest/desktopformfactor) element specifies the settings for an add-in that runs in Office on the web, Windows, and Mac.

The following example shows the `<Hosts>`, `<Host>`, and `<DesktopFormFactor>` elements.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- Information about FunctionFile and ExtensionPoint. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

### Step 3: Add the FunctionFile element

The [`<FunctionFile>` element](/javascript/api/manifest/functionfile) specifies a file that contains JavaScript or TypeScript code to run when an add-in command uses the **ExecuteFunction** action. The `<FunctionFile>` element's **resid** attribute is set to a HTML file that includes all the JavaScript or TypeScript files your add-in commands require. You can't link directly to a JavaScript or TypeScript file. You can only link to an HTML file. The file name is specified as a [`<Url>` element](/javascript/api/manifest/url) in the [`<Resources>` element](/javascript/api/manifest/resources).

> [!NOTE]
> The Yo Office projects use [webpack](https://webpack.js.org/concepts/) to avoid manually adding the JavaScript or TypeScript to the HTML.

The following is an example of the `<FunctionFile>` element.
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="Commands.Url" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- Information about this extension point. -->
    </ExtensionPoint>

    <!-- You can define more than one ExtensionPoint element as needed. -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> Office.js must be initialized before the add-in command logic runs. For more information, see [Initialize your Office Add-in](initialize-add-in.md).

#### Outlook notifications

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

### Step 4: Add ExtensionPoint elements

The [`<ExtensionPoint>` element](/javascript/api/manifest/extensionpoint) defines where add-in commands should appear in the Office UI.

The following examples show how to use the `<ExtensionPoint>` element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.
  
```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso Tab">
  <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element. -->
  <!-- <OfficeTab id="TabData"> -->
    <Label resid="residLabel4" />
    <Group id="Group1Id12">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Control xsi:type="Button" id="Button1Id1">

        <!-- Information about the control. -->
      </Control>
      <!-- Other controls, as needed. -->
    </Group>
  </CustomTab>
</ExtensionPoint>
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="ContextMenu2">
            <!-- Information about the control. -->
    </Control>
    <!-- Other controls, as needed. -->
  </OfficeMenu>
</ExtensionPoint>
```

### Step 5: Add Control elements

The [`<Control>` element](/javascript/api/manifest/control) defines the usable surface of command, such as a button or menu, and the action associated with it.

#### Button controls

A [button control](/javascript/api/manifest/control-button) performs a single action when the user selects it. It can either run a JavaScript or TypeScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the `<Control>` element:

- The **type** attribute is required, and must be set to **Button**.
- The **id** attribute of the `<Control>` element is a string with a maximum of 125 characters.
- The **xsi:type** attribute of the child [`<Action>` element](/javascript/api/manifest/action) must be set to **ExecuteFunction** to run a function or **ShowTaskpane** to display a task pane.

```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
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
    <FunctionName>highlightSelection</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
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

#### Menu controls

A [menu control](/javascript/api/manifest/control-menu) can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:
  
- A root-level menu item.
- A list of submenu items.

When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either run a JavaScript or TypeScript function or show a task pane. Only one level of submenus is supported at this time.

The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the `<Control>` element:

- The **xsi:type** attribute is required, and must be set to **Menu**.
- The **id** attribute is a string with a maximum of 125 characters.

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
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

#### Sample code for function commands

The following code shows a function that's invoked by a button or menu item control whose `<Action>` element's **xsi:type** is set to **ExecuteFunction**. Note the following about the code.

- The [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) call tells Office which function to run when a button or menu item is selected. The value passed to its **actionId** parameter must match the value specified in the [`<FunctionName>` element](/javascript/api/manifest/action#functionname) of the manifest. You must have an `Office.actions.associate` call for every function command defined in the manifest.
- The [event.completed](/javascript/api/office/office.addincommands.event#office-office-addincommands-event-completed-member(1)) call signals that you've successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls `event.completed`, the next queued call to that function runs. You must implement `event.completed`, otherwise your function won't run.

```js
// Initialize the Office Add-in.
Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

// The command function.
async function highlightSelection(event) {

    // Implement your custom code here. The following code is a simple Excel example.
    try {
          await Excel.run(async (context) => {
              const range = context.workbook.getSelectedRange();
              range.format.fill.color = "yellow";
              await context.sync();
          });
      } catch (error) {
          // Note: In a production add-in, notify the user through your add-in's UI.
          console.error(error);
      }

    // Calling event.completed is required. The event.completed call lets the platform know that processing has completed.
    event.completed();
}

// This maps the function to the action ID specified in the manifest.
Office.actions.associate("highlightSelection", highlightSelection);

```

### Step 6: Add the Resources element

The [`<Resources>` element](/javascript/api/manifest/resources) contains resources used by the different child elements of the `<VersionOverrides>` element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.
  
The following shows an example of how to use the `<Resources>` element. Each resource can have one or more [`<Override>` child elements](/javascript/api/manifest/override) to define a different resource for a specific locale.

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

> [!NOTE]
> You must use Secure Sockets Layer (SSL) for all URLs in the `<Image>` and `<Url>` elements.

## Outlook support notes

Add-in commands are available in the following Outlook versions.

- Outlook on the web for Microsoft 365 and Outlook.com
- Outlook on the web for Exchange 2016 or later
- [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)
- Outlook 2016 or later on Windows
- Outlook on Mac
- Outlook on Android
- Outlook on iOS

Support for add-in commands in Exchange 2016 requires [Cumulative Update 5](https://support.microsoft.com/topic/d67d7693-96a4-fb6e-b60b-e64984e267bd).

If your add-in uses an add-in only manifest, then add-in commands are only available for add-ins that don't use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](/javascript/api/manifest/rule) to limit the types of items they activate on. However, [contextual add-ins](../outlook/contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).

## See also

- [Add-in commands](../design/add-in-commands.md)
- [Sample: Create an Excel add-in with command buttons](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [Sample: Create a Word add-in with command buttons](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [Sample: Create a PowerPoint add-in with command buttons](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
