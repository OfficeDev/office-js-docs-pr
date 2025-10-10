---
title: Office Add-ins with the add-in only manifest
description: Get an overview of the add-in only manifest for Office add-ins and its uses.
ms.topic: overview
ms.date: 06/24/2025
ms.localizationpriority: high
---

# Office Add-ins with the add-in only manifest

This article introduces the XML-formatted add-in only manifest for Office Add-ins. It assumes that you're familiar with the [Office Add-ins manifest](add-in-manifests.md).

> [!TIP]
> For an overview of the unified manifest for Microsoft 365, see [Office Add-ins with the unified manifest for Microsoft 365](unified-manifest-overview.md).

## Schema versions

Not all Office clients support the latest features, and some Office users will have an older version of Office. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

The `<VersionOverrides>` element in the manifest is an example of this. All elements defined inside `<VersionOverrides>` will override the same element in the other part of the manifest. This means that, whenever possible, Office will use what is in the `<VersionOverrides>` section to set up the add-in. However, if the version of Office doesn't support a certain version of `<VersionOverrides>`, Office will ignore it and depend on the information in the rest of the manifest.

This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.

The current versions of the schema are:

|Version|Description|
|:-----|:-----|
|v1.0|Supports version 1.0 of the Office JavaScript API. For example, in Outlook add-ins, this supports the read form. |
|v1.1|Supports version 1.1 of the Office JavaScript API and `<VersionOverrides>`. For example, in Outlook add-ins, this adds support for the compose form.|
|`<VersionOverrides>` 1.0|Supports later versions of the Office JavaScript API. This supports add-in commands.|
|`<VersionOverrides>` 1.1|Supported by Outlook only. This version of `<VersionOverrides>` adds support for newer features, such as [pinnable task panes](../outlook/pinnable-taskpane.md) and mobile add-ins.|

Even if your add-in manifest uses the `<VersionOverrides>` element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support `<VersionOverrides>`.

> [!NOTE]
> Office uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. See [How to find the proper order of manifest elements](../develop/manifest-element-ordering.md) elements in the required order.

## Required elements

The following table specifies the elements that are required for the three types of Office Add-ins.

> [!NOTE]
> There is also a mandatory order in which elements must appear within their parent element. For more information see [How to find the proper order of add-in only manifest elements](manifest-element-ordering.md).

### Required elements by Office Add-in type

| Element                                                                                      | Content    | Task pane    | Mail<br>(Outlook) |
| :------------------------------------------------------------------------------------------- | :--------: | :----------: | :--------:   |
| [OfficeApp][]                                                                                | Required   | Required     | Required     |
| [Id][]                                                                                       | Required   | Required     | Required     |
| [Version][]                                                                                  | Required   | Required     | Required     |
| [ProviderName][]                                                                             | Required   | Required     | Required     |
| [DefaultLocale][]                                                                            | Required   | Required     | Required     |
| [DisplayName][]                                                                              | Required   | Required     | Required     |
| [Description][]                                                                              | Required   | Required     | Required     |
| [SupportUrl][]\*\*                                                                           | Required   | Required     | Required     |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       | Required   | Required     | Not available|
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]<br/>[SourceLocation (MailApp)][]| Required | Required | Required   |
| [DesktopSettings][]                                                                          | Not available | Not available | Required |
| [Permissions (ContentApp)][]<br/>[Permissions (TaskPaneApp)][]<br/>[Permissions (MailApp)][] | Required   | Required     | Required     |
| [Rule (RuleCollection)][]<br/>[Rule (MailApp)][]                                             | Not available | Not available | Required |
| [Requirements (MailApp)][]\*                                                                 | Not applicable| Not available | Required |
| [Set][]\*<br/>[Sets (Requirements)][]\*<br/>[Sets (MailAppRequirements)][]\*                 | Required   | Required     | Required     |
| [Form][]\*<br/>[FormSettings][]\*                                                            | Not available | Not available | Required |
| [Hosts][]\*                                                                                  | Required   | Required     | Optional     |

_\*Added in the Office Add-in Manifest Schema version 1.1._

_\*\* SupportUrl is only required for add-ins that are distributed through Microsoft Marketplace._

<!-- Links for above table -->

[officeapp]: /javascript/api/manifest/officeapp
[id]: /javascript/api/manifest/id
[version]: /javascript/api/manifest/version
[providername]: /javascript/api/manifest/providername
[defaultlocale]: /javascript/api/manifest/defaultlocale
[displayname]: /javascript/api/manifest/displayname
[description]: /javascript/api/manifest/description
[supporturl]: /javascript/api/manifest/supporturl
[defaultsettings (contentapp)]: /javascript/api/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /javascript/api/manifest/defaultsettings
[sourcelocation (contentapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (mailapp)]: /javascript/api/manifest/sourcelocation
[desktopsettings]: /javascript/api/manifest/desktopsettings
[permissions (contentapp)]: /javascript/api/manifest/permissions
[permissions (taskpaneapp)]: /javascript/api/manifest/permissions
[permissions (mailapp)]: /javascript/api/manifest/permissions
[rule (rulecollection)]: /javascript/api/manifest/rule
[rule (mailapp)]: /javascript/api/manifest/rule
[requirements (mailapp)]: /javascript/api/manifest/requirements
[set]: /javascript/api/manifest/set
[sets (mailapprequirements)]: /javascript/api/manifest/sets
[form]: /javascript/api/manifest/form
[formsettings]: /javascript/api/manifest/formsettings
[sets (requirements)]: /javascript/api/manifest/sets
[hosts]: /javascript/api/manifest/hosts

## Root element

The root element for the Office Add-in manifest is `<OfficeApp>`. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element.

```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- The rest of the manifest. -->

</OfficeApp>
```

## Version

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will be unable to use the add-in until consent is granted.

## Hosts

Office add-ins specify the `<Hosts>` element like the following:

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

This is separate from the `<Hosts>` element inside the `<VersionOverrides>` element, which is discussed in [Create add-in commands with the add-in only manifest](../develop/create-addin-commands.md).

## Specify safe domains with the AppDomains element

There is an [AppDomains](/javascript/api/manifest/appdomains) element of the add-in only manifest file that is used to tell Office which domains your add-in should be allowed to navigate to. As noted in [Specify domains you want to open in the add-in window](add-in-manifests.md#specify-domains-you-want-to-open-in-the-add-in-window), when running in Office on the web, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](/javascript/api/manifest/sourcelocation) element), that URL opens in a new browser window outside the add-in pane of the Office application.

To override this (desktop Office) behavior, add each domain you want to open in the add-in window in the list of domains specified in the `<AppDomains>` element. If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop. If it tries to go to a URL that isn't in the list, then in desktop Office that URL opens in a new browser window (outside the add-in pane).

The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.

|Office client|Domain defined in AppDomains?|Browser behavior|
|---|---|---|
|All clients|Yes|Link opens in add-in task pane.|
|Office 2016 on Windows (volume-licensed perpetual)|No|Link opens in Internet Explorer 11.|
|Other clients|No|Link opens in user's default browser.|

The following add-in only manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the `<SourceLocation>` element. It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](/javascript/api/manifest/appdomain) element within the `<AppDomains>` element list. If the add-in goes to a page in the `www.northwindtraders.com` domain, that page opens in the add-in pane, even in Office desktop.

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## Version overrides in the manifest

The optional [VersionOverrides](/javascript/api/manifest/versionoverrides) element contains child markup that enables additional add-in features. Some of these are:

- Customizing the Office ribbon and menus.
- Customizing how Office works with the embedded runtimes in which add-ins run.
- Configuring how the add-in interacts with Microsoft Entra ID and Microsoft Graph for Single Sign-on.

Some descendant elements of `VersionOverrides` have values that override values of the parent `OfficeApp` element. For example, the `Hosts` element in `VersionOverrides` overrides the `Hosts` element in `OfficeApp`.

The `VersionOverrides` element has its own schema, actually four of them, depending on the type of add-in and the features it uses. The schemas are:

- [Task pane 1.0](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)
- [Content 1.0](/openspecs/office_file_formats/ms-owemxml/c9cb8dca-e9e7-45a7-86b7-f1f0833ce2c7)
- [Mail 1.0](/openspecs/office_file_formats/ms-owemxml/578d8214-2657-4e6a-8485-25899e772fac)
- [Mail 1.1](/openspecs/office_file_formats/ms-owemxml/8e722c85-eb78-438c-94a4-edac7e9c533a)

When a `VersionOverrides` element is used, then the `OfficeApp` element must have a `xmlns` attribute that identifies the appropriate schema. The possible values of the attribute are the following:

- `http://schemas.microsoft.com/office/taskpaneappversionoverrides`
- `http://schemas.microsoft.com/office/contentappversionoverrides`
- `http://schemas.microsoft.com/office/mailappversionoverrides`

The `VersionOverrides` element itself must also have an `xmlns` attribute specifying the schema. The possible values are the three above and the following:

- `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`

The `VersionOverrides` element also must have an `xsi:type` attribute that specifies the schema version. The possible values are the following:

- `VersionOverridesV1_0`
- `VersionOverridesV1_1`

The following are examples of `VersionOverrides` used, respectively, in a task pane add-in and a mail add-in. Note that when a mail `VersionOverrides` with version 1.1 is used, it must be the last child of a parent `VersionOverrides` of type 1.0. The values of child elements in the inner `VersionOverrides` override the values of the same-named elements in the parent `VersionOverrides` and the grandparent `OfficeApp` element.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <!-- Child elements are omitted. -->
</VersionOverrides>
```

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <!-- Other child elements are omitted. -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <!-- Child elements are omitted. -->
  </VersionOverrides>
</VersionOverrides>
```

For an example of a manifest that includes a `VersionOverrides` element, see [Manifest v1.1 XML file examples and schemas](#manifest-v11-xml-file-examples-and-schemas).

## Requirements

The `<Requirements>` element specifies the set of APIs available to the add-in. For detailed information about requirement sets, see [Office requirement sets availability](office-versions-and-requirement-sets.md#office-requirement-sets-availability). For example, in an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above.

The `<Requirements>` element can also appear in the `<VersionOverrides>` element, allowing the add-in to specify a different requirement when loaded in clients that support `<VersionOverrides>`.

The following example uses the **DefaultMinVersion** attribute of the `<Sets>` element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the `<Set>` element to require the Mailbox requirement set version 1.1.

```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```

## Localization

Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the `<Resources>` element within the `<VersionOverrides>` element. The following shows how to override an image, a URL, and a string.

```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- Add information for other locales. -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- Add information for other locales. -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- Add information for other locales. -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

The [schema reference](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) contains full information on which elements can be localized.

## Manifest v1.1 XML file examples and schemas

The following sections show examples of manifest v1.1 XML files for content, task pane, and mail (Outlook) add-ins.

# [Task pane](#tab/tabid-1)

[Add-in manifest schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation. -->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided. -->

  <!-- IMPORTANT! Id must be unique for your add-in. If you copy this manifest, ensure that you change this ID to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-in's dialog. -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions." />
  <!-- Icon for your add-in. Used on installation screens and the add-in's dialog. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!-- Domains that are allowed when navigating. For example, if you use ShowTaskpane and then have an href link, navigation is only allowed if the domain is on this list. -->
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
  </AppDomains>
  <!-- End Basic Settings. -->

  <!-- BeginTaskPaneMode integration. Any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides. -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!-- EndTaskPaneMode integration. -->

  <Permissions>ReadWriteDocument</Permissions>

  <!-- BeginAddinCommandsMode integration. -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!-- Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel, Document=Word, Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest. -->
      <Host xsi:type="Document">
        <!-- Form factor. DesktopFormFactor is supported. Other form factors are available depending on the host and feature. -->
        <DesktopFormFactor>
          <!-- This code enables a customizable message to be displayed when the add-in is loaded successfully upon individual install. -->
          <GetStarted>
            <!-- Title of the Getting Started callout. The resid attribute points to a ShortString resource. -->
            <Title resid="Contoso.GetStarted.Title"/>
            <!-- Description of the Getting Started callout. resid points to a LongString resource. -->
            <Description resid="Contoso.GetStarted.Description"/>  
            <!-- Points to a URL resource which details how the add-in should be used. -->
            <LearnMoreUrl resid="Contoso.GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <!-- Function file is an HTML page that includes, or loads, the JavaScript where functions for ExecuteAction will be called. Think of the FunctionFile as the "code behind" ExecuteFunction. -->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!-- PrimaryCommandSurface==Main Office app ribbon. -->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!-- Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab. -->
            <!-- Documentation includes all the IDs currently tested to work. -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource. -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Each size needs its own icon resource or it will look distorted when resized. -->
                  <!-- Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!-- Use PNG icons and remember that all URLs on the resources section must use HTTPS. -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!-- Control. It can be of type "Button" or "Menu". -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!-- Label for your button. resid must point to a ShortString resource. -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!-- ToolTip title. resid must point to a ShortString resource. -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!-- ToolTip description. resid must point to a LongString resource. -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!-- This is what happens when the command is triggered (e.g., click on the ribbon button). -->
                  <!-- Supported actions are ExecuteFunction or ShowTaskpane. -->
                  <!-- Look at the FunctionFile.html page for reference on how to implement the function. -->
                  <Action xsi:type="ExecuteFunction">
                    <!-- Name of the function to call. This function needs to exist in the global DOM namespace of the function file. -->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!-- Provide a URL resource ID for the location that will be displayed on the task pane. -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example. -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab. -->
              <!-- If validating with XSD, it needs to be at the end. -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <!-- You can use resources across hosts and form factors. -->
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://myCDN/Images/ButtonFunction.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <!-- ShortStrings max characters=125. -->
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Task Pane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Task Pane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Task Pane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <!-- LongStrings max characters=250. -->
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to execute function." />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to show a task pane." />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to show options on this menu." />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to show Task Pane 1." />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to show Task Pane 2." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
  <!-- EndAddinCommandsMode integration. -->
</OfficeApp>
```

# [Content](#tab/tabid-2)

[Add-in manifest schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# [Mail](#tab/tabid-3)

[Add-in manifest schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" xsi:type="MailApp">

  <!-- Begin basic settings: Add-in metadata used for all versions of Outlook, unless override provided. -->

  <!-- IMPORTANT: The ID must be unique to your add-in. If you reuse this manifest, ensure that you change this to a new GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052C8</Id>
  <!-- Updates from the Office Store only get triggered if there is a version change. -->
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used in the in-app Office Store and various places of the Outlook UI, such as an add-in's dialog. -->
  <DisplayName DefaultValue="Contoso Add-in"/>
  <Description DefaultValue="An Outlook add-in template to get started."/>
  <!-- Change the following lines to specify the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png"/>
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]"/>
  <!-- Domains allowed for navigation. -->
  <AppDomains>
    <AppDomain>https://www.contoso.com</AppDomain>
  </AppDomains>

  <!--End basic settings. -->

  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <!-- The <Requirements> element is overridden by any <Requirements> element inside a <VersionOverrides> element. -->
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <!-- The <FormSettings> element is required for validation, but is ignored when there's a <VersionOverrides> element in your manifest. -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue="[Insert the URL where your HTML file is hosted.]"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <!-- The <Rule> element is required for validation, but is ignored when there's a <VersionOverrides> element in your manifest. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.13">
          <bt:Set Name="Mailbox"/>
        </bt:Sets>
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <!-- Location of the functions that will run when the add-in's function command is selected. -->
            <FunctionFile resid="functionFile"/>
            <!-- Activates the add-in on the Message Read surface. -->
            <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <!-- Use the default tab of the ExtensionPoint or create your own with <CustomTab id="myTab">. -->
              <OfficeTab id="TabDefault">
                <!-- Add up to six groups per tab. -->
                <Group id="msgReadGroup">
                  <Label resid="groupLabel"/>
                  <!-- Configures the button to launch the add-in's task pane. -->
                  <Control xsi:type="Button" id="msgReadOpenPaneButton">
                    <Label resid="taskPaneButtonLabel"/>
                    <Supertip>
                      <Title resid="taskPaneButtonLabel"/>
                      <Description resid="taskPaneButtonDescription"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="icon16"/>
                      <bt:Image size="32" resid="icon32"/>
                      <bt:Image size="80" resid="icon80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="messageReadTaskPaneUrl"/>
                    </Action>
                  </Control>
                  <!-- Configures the function command of the add-in. -->
                  <Control xsi:type="Button" id="msgReadActionButton">
                    <Label resid="actionButtonLabel"/>
                    <Supertip>
                      <Title resid="actionButtonLabel"/>
                      <Description resid="actionButtonDescription"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="icon16"/>
                      <bt:Image size="32" resid="icon32"/>
                      <bt:Image size="80" resid="icon80"/>
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>run</FunctionName>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
      <!-- You can use resources across hosts and form factors. -->
      <Resources>
        <bt:Images>
          <bt:Image id="icon16" DefaultValue="https://contoso.com/assets/icon16.png"/>
          <bt:Image id="icon32" DefaultValue="https://contoso.com/assets/icon32.png"/>
          <bt:Image id="icon80" DefaultValue="https://contoso.com/assets/icon80.png"/>
        </bt:Images>
        <bt:Urls>
          <bt:Url id="functionFile" DefaultValue="https://contoso.com/FunctionFile.html"/>
          <bt:Url id="messageReadTaskPaneUrl" DefaultValue="https://contoso.com/MessageRead.html"/>
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="groupLabel" DefaultValue="My Add-in Group"/>
          <bt:String id="taskPaneButtonLabel" DefaultValue="Show Task Pane"/>
          <bt:String id="actionButtonLabel" DefaultValue="Perform an Action"/>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="taskPaneButtonDescription" DefaultValue="Opens a task pane."/>
          <bt:String id="actionButtonDescription" DefaultValue="Performs an action."/>
        </bt:LongStrings>
      </Resources>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
```

---

## Validate an Office Add-in's manifest

For information about validating a manifest against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).

## See also

- [How to find the proper order of add-in only manifest elements](manifest-element-ordering.md)
- [Create add-in commands with the add-in only manifest](create-addin-commands.md)
- [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md)
- [Localization for Office Add-ins](localization.md)
- [Schema reference for XML Office Add-ins manifests](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
- [Make your Office Add-in compatible with an existing COM or VSTO add-in](make-office-add-in-compatible-with-existing-com-add-in.md)
- [Requesting permissions for API use in add-ins](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md)
