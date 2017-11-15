---
title: Office Add-ins XML manifest
description: 
ms.date: 11/20/2017 
---


# Office Add-ins XML manifest


The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications. 

An XML manifest file based on this schema enables an Office Add-in to do the following:

- Describe itself by providing an ID, version, description, display name, and default locale.
    
- Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.
    
- Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.
    
- Declare permissions that the Office Add-in requires, such as reading or writing to the document.
    
- For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.

> [!NOTE]
> When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to the Office Store, make sure that you conform to the [Office Store validation policies](https://msdn.microsoft.com/en-us/library/jj220035.aspx). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) and the [Office Add-in host and availability page](https://dev.office.com/add-in-availability)).


## Required elements


The following table specifies the elements that are required for the three types of Office Add-ins.


**Required elements by Office Add-in type**


|**Element**|**Content**|**Task pane**|**Outlook**|
|:-----|:-----|:-----|:-----|
|[OfficeApp](http://msdn.microsoft.com/en-us/library/68f1cada-66f8-4341-45f5-14e0634c24fb%28Office.15%29.aspx)|X|X|X|
|[Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx)|X|X|X|
|[Version](http://msdn.microsoft.com/en-us/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx)|X|X|X|
|[ProviderName](http://msdn.microsoft.com/en-us/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx)|X|X|X|
|[DefaultLocale](http://msdn.microsoft.com/en-us/library/04796a3a-3afa-dc85-db66-4677560c185c%28Office.15%29.aspx)|X|X|X|
|[DisplayName](http://msdn.microsoft.com/en-us/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx)|X|X|X|
|[Description](http://msdn.microsoft.com/en-us/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx)|X|X|X|
|[IconUrl](http://msdn.microsoft.com/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx)|X|X|X|
|[HighResolutionIconUrl](http://msdn.microsoft.com/library/ff7b2647-ec8e-70dc-4e4a-e1a1377ff3f2%28Office.15%29.aspx)|X|X|X|
|[DefaultSettings (ContentApp)](http://msdn.microsoft.com/en-us/library/f7edc689-551f-1a17-ea81-ffd58f534557%28Office.15%29.aspx)<br/>[DefaultSettings (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/36e3d139-56a4-fb3d-0a21-cbd14e606765%28Office.15%29.aspx)|X|X||
|[SourceLocation (ContentApp)](http://msdn.microsoft.com/en-us/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx)<br/>[SourceLocation (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx)|X|X||
|[DesktopSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx)|||X|
|[SourceLocation (MailApp)](http://msdn.microsoft.com/en-us/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx)|||X|
|[Permissions (ContentApp)](http://msdn.microsoft.com/en-us/library/9f3dcf9c-fced-c115-4f0d-38d60fb7c583%28Office.15%29.aspx)<br/>[Permissions (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx)<br/>[Permissions (MailApp)](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx)|X|X|X|
|[Rule (RuleCollection)](http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx)<br/>[Rule (MailApp)](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx)|||X|
|[*Requirements (MailApp)](http://msdn.microsoft.com/en-us/library/9536ea30-34f7-76b5-7f30-1508626840e4%28Office.15%29.aspx)|||X|
|[*Set](http://msdn.microsoft.com/en-us/library/1506daa1-332c-30e1-6402-3371bcd0b895%28Office.15%29.aspx)<br/>[**Sets (MailAppRequirements)](http://msdn.microsoft.com/en-us/library/2a6a2484-eeee-37e4-43bc-c185e8ae0d1d%28Office.15%29.aspx)|||X|
|[*Form](http://msdn.microsoft.com/en-us/library/77a8ac83-c22b-1225-4fc4-ba4038b68648%28Office.15%29.aspx)<br/>[**FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx)|||X|
|[*Sets (Requirements)](http://msdn.microsoft.com/en-us/library/509be287-b532-87c6-71ac-64f3a4bbd3af%28Office.15%29.aspx)|X|X||
|[*Hosts](http://msdn.microsoft.com/library/f9a739c1-3daf-c03a-2bd9-4a2a6b870101%28Office.15%29.aspx)|X|X||

*Added in the Office Add-in Manifest Schema version 1.1.

> [!IMPORTANT]
> - Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID. 

> - All URLs, such as the source file locations specified in the [SourceLocation](../../reference/manifest/sourcelocation.md) element, must be **SSL-secured (HTTPS)**.

> - All icon URLs, such as those used on command surfaces, must **allow caching**. The web server should NOT return HTTP headers like no-cache/no-store. 

> - Add-ins submitted to the Office Store must also include the [SupportUrl](../../reference/manifest/supporturl.md) element. For more information, see [What are some common submission errors to avoid?](http://msdn.microsoft.com/library/0ceb385c-a608-40cc-8314-78e39d6c75d0%28Office.15%29.aspx#bk_q2)

> - As a best practice, only use the [AppDomains](../../reference/manifest/appdomains.md) element to specify domains other than the one specified in the [SourceLocation](../../reference/manifest/sourcelocation.md) element for authentication scenarios.

## Specify domains you want to open in the add-in window

By default, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](../../reference/manifest/sourcelocation.md) element of the manifest file), that URL will open in a new browser window outside the add-in pane of the Office host application. This default behavior protects the user against unexpected page navigation within the add-in pane from embedded **iframe** elements.

To override this behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](../../reference/manifest/appdomains.md) element of the manifest file. If the add-in tries to go to a URL in a domain that isn't in the list, that URL will open in a new browser window (outside the add-in pane).

The following XML manifest example hosts its main add-in page in the  `https://www.contoso.com` domain as specified in the **SourceLocation** element. It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](http://msdn.microsoft.com/en-us/library/2a0353ec-5e09-6fbf-1636-4bb5dcebb9bf%28Office.15%29.aspx) element within the **AppDomains** element list. If the add-in goes to a page in the www.northwindtraders.com domain, that page will open in the add-in pane.


```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## Manifest v1.1 XML file examples and schemas


The following sections show examples of manifest v1.1 XML files for content, task pane,  and Outlook add-ins.

### Office Add-in manifest v1.1 example with commands and fallback task pane
[Task pane manifest schema](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

<!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

<!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

<!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
   <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
   <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

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
      	<!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
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
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
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
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
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

### Content add-in manifest v1.1 example
[Content manifest schema](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas/content)


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
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

### Outlook add-in manifest v1.1 example
[Content manifest schema](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas/mail)


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="MailApp">

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
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

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

## Validate and troubleshoot issues with your manifest

For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../../docs/testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.

## Additional resources


- [Define add-in commands in your manifest](../../docs/develop/define-add-in-commands.md)
- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Localization for Office Add-ins](../../docs/develop/localization.md)
- [Schema reference for Office Add-ins manifests](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas)
- [Validate and troubleshoot issues with your manifest](../../docs/testing/troubleshoot-manifest.md)

