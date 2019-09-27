---
title: Office Add-ins XML manifest
description: ''
ms.date: 09/26/2019
localization_priority: Priority
---

# Office Add-ins XML manifest

The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.

An XML manifest file based on this schema enables an Office Add-in to do the following:

* Describe itself by providing an ID, version, description, display name, and default locale.

* Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office Ribbon.

* Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.

* Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.

* Declare permissions that the Office Add-in requires, such as reading or writing to the document.

* For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.

> [!NOTE]
> If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).

## Required elements

The following table specifies the elements that are required for the three types of Office Add-ins.

> [!NOTE]
> There is also a mandatory order in which elements must appear within their parent element. For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).


### Required elements by Office Add-in type

| Element                                                                                      | Content | Task pane | Outlook |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| [OfficeApp][]                                                                                |    X    |     X     |    X    |
| [Id][]                                                                                       |    X    |     X     |    X    |
| [Version][]                                                                                  |    X    |     X     |    X    |
| [ProviderName][]                                                                             |    X    |     X     |    X    |
| [DefaultLocale][]                                                                            |    X    |     X     |    X    |
| [DisplayName][]                                                                              |    X    |     X     |    X    |
| [Description][]                                                                              |    X    |     X     |    X    |
| [IconUrl][]                                                                                  |    X    |     X     |    X    |
| [SupportUrl][]\*\*                                                                           |    X    |     X     |    X    |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       |    X    |     X     |         |
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]                         |    X    |     X     |         |
| [DesktopSettings][]                                                                          |         |           |    X    |
| [SourceLocation (MailApp)][]                                                                 |         |           |    X    |
| [Permissions (ContentApp)][]<br/>[Permissions (TaskPaneApp)][]<br/>[Permissions (MailApp)][] |    X    |     X     |    X    |
| [Rule (RuleCollection)][]<br/>[Rule (MailApp)][]                                             |         |           |    X    |
| [Requirements (MailApp)*][]                                                                  |         |           |    X    |
| [Set*][]<br/>[Sets (MailAppRequirements)*][]                                                 |         |           |    X    |
| [Form*][]<br/>[FormSettings*][]                                                              |         |           |    X    |
| [Sets (Requirements)*][]                                                                     |    X    |     X     |         |
| [Hosts*][]                                                                                   |    X    |     X     |         |

_\*Added in the Office Add-in Manifest Schema version 1.1._

_\*\* SupportUrl is only required for add-ins that are distributed through AppSource._

<!-- Links for above table -->

[officeapp]: /office/dev/add-ins/reference/manifest/officeapp
[id]: /office/dev/add-ins/reference/manifest/id
[version]: /office/dev/add-ins/reference/manifest/version
[providername]: /office/dev/add-ins/reference/manifest/providername
[defaultlocale]: /office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: /office/dev/add-ins/reference/manifest/displayname
[description]: /office/dev/add-ins/reference/manifest/description
[iconurl]: /office/dev/add-ins/reference/manifest/iconurl
[supporturl]: /office/dev/add-ins/reference/manifest/supporturl
[defaultsettings (contentapp)]: /office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: /office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: https://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissions (contentapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (taskpaneapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (mailapp)]: /office/dev/add-ins/reference/manifest/permissions
[rule (rulecollection)]: /office/dev/add-ins/reference/manifest/rule
[rule (mailapp)]: /office/dev/add-ins/reference/manifest/rule
[requirements (mailapp)*]: /office/dev/add-ins/reference/manifest/requirements
[set*]: /office/dev/add-ins/reference/manifest/set
[sets (mailapprequirements)*]: /office/dev/add-ins/reference/manifest/sets
[form*]: /office/dev/add-ins/reference/manifest/form
[formsettings*]: /office/dev/add-ins/reference/manifest/formsettings
[sets (requirements)*]: /office/dev/add-ins/reference/manifest/sets
[hosts*]: /office/dev/add-ins/reference/manifest/hosts

## Hosting requirements

All image URIs, such as those used for [add-in commands][], must support caching. The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.

All URLs, such as the source file locations specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**. [!include[HTTPS guidance](../includes/https-guidance.md)]

## Best practices for submitting to AppSource

Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.

Add-ins submitted to AppSource must also include the [SupportUrl](/office/dev/add-ins/reference/manifest/supporturl) element. For more information, see [Validation policies for apps and add-ins submitted to AppSource](/office/dev/store/validation-policies).

Only use the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element for authentication scenarios.

## Specify domains you want to open in the add-in window

When running in Office on the web, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office host application.

To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file. If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop. If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).

> [!NOTE]
> There are two exceptions to this behavior:
> 
> - It applies only to the root pane of the add-in. If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.
> - When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office. 

The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element. It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](/office/dev/add-ins/reference/manifest/appdomain) element within the **AppDomains** element list. If the add-in goes to a page in the www.northwindtraders.com domain, that page opens in the add-in pane, even in Office desktop.

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

## Specify domains from which Office.js API calls are made

Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file. If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file. If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md). 

## Manifest v1.1 XML file examples and schemas

The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.

# [Task pane](#tab/tabid-1)

[Task pane app manifest schema](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# [Content](#tab/tabid-2)

[Content app manifest schema](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

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

[Mail app manifest schema](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## Validate and troubleshoot issues with your manifest

For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.

## See also

* [Create add-in commands in your manifest][add-in commands]
* [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md)
* [Localization for Office Add-ins](localization.md)
* [Schema reference for Office Add-ins manifests](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md)

[add-in commands]: create-addin-commands.md
