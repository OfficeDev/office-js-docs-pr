---
title: Outlook add-in manifests
description: Get an overview of the two kinds of manifests available for Outlook Add-ins.
ms.date: 01/23/2023
ms.localizationpriority: high
---

# Outlook add-in manifests

An Outlook add-in consists of two components: the add-in manifest and a web app supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients.

There are two possible formats for the manifest: XML and JSON. You can learn about the JSON manifest at [Unified Microsoft 365 manifest (preview)](../develop/json-manifest-overview.md). This article is about the XML manifest.

The following is an example of the XML manifest.

 > [!NOTE]
 > All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="https://appdemo.contoso.com/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="https://appdemo.contoso.com/images/blue-128.png" />
  <SupportUrl DefaultValue="https://appdemo.contoso.com"/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read task pane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />

          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <!-- Function (UI-less) button -->
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
                <!-- Menu (dropdown) button -->
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
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
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
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="https://appdemo.contoso.com/images/blue-16.png" />
        <bt:Image id="blue-icon-32" DefaultValue="https://appdemo.contoso.com/images/blue-32.png" />
        <bt:Image id="blue-icon-80" DefaultValue="https://appdemo.contoso.com/images/blue-80.png" />
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="https://appdemo.contoso.com/images/red-16.png" />
        <bt:Image id="red-icon-32" DefaultValue="https://appdemo.contoso.com/images/red-32.png" />
        <bt:Image id="red-icon-80" DefaultValue="https://appdemo.contoso.com/images/red-80.png" />
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="https://appdemo.contoso.com/images/green-16.png" />
        <bt:Image id="green-icon-32" DefaultValue="https://appdemo.contoso.com/images/green-32.png" />
        <bt:Image id="green-icon-80" DefaultValue="https://appdemo.contoso.com/images/green-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://appdemo.contoso.com/FunctionFile/Functions.html" />
        <bt:Url id="readTaskPaneUrl" DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## Schema versions

Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

The **\<VersionOverrides\>** element in the manifest is an example of this. All elements defined inside **\<VersionOverrides\>** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **\<VersionOverrides\>** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **\<VersionOverrides\>**, Outlook will ignore it and depend on the information in the rest of the manifest.

This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.

The current versions of the schema are:

|Version|Description|
|:-----|:-----|
|v1.0|Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form. |
|v1.1|Supports version 1.1 of the Office JavaScript API and **\<VersionOverrides\>**. For Outlook add-ins, this adds support for compose form.|
|**\<VersionOverrides\>** 1.0|Supports later versions of the Office JavaScript API. This supports add-in commands.|
|**\<VersionOverrides\>** 1.1|Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.|

This article will cover the requirements for a v1.1 manifest. Even if your add-in manifest uses the **\<VersionOverrides\>** element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support **\<VersionOverrides\>**.

> [!NOTE]
> Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.

## Root element

The root element for the Outlook add-in manifest is **\<OfficeApp\>**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element.

```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest -->

</OfficeApp>
```

## Version

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will be unable to use the add-in until consent is granted.

## VersionOverrides

The **\<VersionOverrides\>** element is the location of information for [add-in commands](../design/add-in-commands.md).

This element is also where add-ins define support for [mobile add-ins](add-mobile-support.md).

## Localization

Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the **\<Resources\>** element within the **\<VersionOverrides\>** element. The following shows how to override an image, a URL, and a string.

```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- add information for other locales -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- add information for other locales -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- add information for other locales -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

The schema reference contains full information on which elements can be localized.

## Hosts

Outlook add-ins specify the **\<Hosts\>** element like the following:

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

This is separate from the **\<Hosts\>** element inside the **\<VersionOverrides\>** element, which is discussed in [Create add-in commands with the XML manifest](../develop/create-addin-commands.md).

## Requirements

The **\<Requirements\>** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](apis.md) for more information on requirement sets.

The **\<Requirements\>** element can also appear in the **\<VersionOverrides\>** element, allowing the add-in to specify a different requirement when loaded in clients that support **\<VersionOverrides\>**.

The following example uses the **DefaultMinVersion** attribute of the **\<Sets\>** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **\<Set\>** element to require the Mailbox requirement set version 1.1.

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

## Form settings

The **\<FormSettings\>** element is used by older Outlook clients, which only support schema 1.1 and not **\<VersionOverrides\>**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.

These settings are directly related to the activation rules in the **\<Rule\>** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.

For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).

## App domains

The domain of the add-in start page that you specify in the **\<SourceLocation\>** element is the default domain for the add-in. Without using the **\<AppDomains\>** and **\<AppDomain\>** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **\<AppDomains\>** element and include each additional domain in its own **\<AppDomain\>** sub-element in the add-in manifest.

The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane.

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

App domains are also necessary to enable cookie sharing between the pop-out window and the add-in running in the rich client.

The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.

|Outlook client|Domain defined<br>in AppDomains?|Browser behavior|
|---|---|---|
|All clients|Yes|Link opens in add-in task pane.|
|- Outlook 2016 on Windows (volume-licensed perpetual)<br>- Outlook 2013 on Windows (perpetual)|No|Link opens in Internet Explorer 11.|
|Other clients|No|Link opens in user's default browser.|

For more details, see the [Specify domains you want to open in the add-in window](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).

## Permissions

The **\<Permissions\>** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), and does not call [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).

**Four-tier permissions model for mail add-ins**

![4-tier permissions model for mail apps schema v1.1.](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## Activation rules

Activation rules are specified in the **\<Rule\>** element. The **\<Rule\>** element can appear as a child of the **\<OfficeApp\>** element in 1.1 manifests.

Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.

> [!NOTE]
> Activation rules only apply to clients that do not support the **\<VersionOverrides\>** element.

- The item type and/or message class

- The presence of a specific type of known entity, such as an address or phone number

- A regular expression match in the body, subject, or sender email address

- The presence of an attachment

For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).

## Next steps: add-in commands

After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands](../design/add-in-commands.md).

For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).

## Next steps: Add mobile support

Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and on Mac. For more information, see [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).

## See also

- [Localization for Office Add-ins](../develop/localization.md)
- [Privacy, permissions, and security for Outlook add-ins](privacy-and-security.md)
- [Outlook add-in APIs](apis.md)
- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md)
- [Design your Office Add-ins](../design/add-in-design.md)
- [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md)
- [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md)
