---
title: Module extension Outlook add-ins
description: Create applications that run inside Outlook to make it easy for your users to access business information and productivity tools without ever leaving Outlook.
ms.date: 08/30/2022
ms.localizationpriority: medium
---

# Module extension Outlook add-ins

Module extension add-ins appear in the Outlook navigation bar, right alongside mail, tasks, and calendars. A module extension is not limited to using mail and appointment information. You can create applications that run inside Outlook to make it easy for your users to access business information and productivity tools without ever leaving Outlook.

> [!TIP]
> Module extensions aren't supported in the [Teams manifest (preview)](../develop/json-manifest-overview.md), but you can create a very similar experience for users by making a [personal tab that opens in Outlook](/microsoftteams/platform/m365-apps/extend-m365-teams-personal-tab). In the early preview period for the Teams manifest in Outlook Add-ins, it isn't possible to combine an Outlook Add-in and a personal tab in the same manifest and install them as a unit. We're working on this, but in the meantime, you must create separate apps for the add-in and the personal tab. They can both use files on the same domain.

> [!NOTE]
> Module extensions are only supported by Outlook 2016 or later on Windows.  

## Open a module extension

To open a module extension, users click on the module's name or icon in the Outlook navigation bar. If the user has compact navigation selected, the navigation bar has an icon that shows an extension is loaded.

![Shows the compact navigation bar when a module extension is loaded in Outlook.](../images/outlook-module-navigationbar-compact.png)

If the user is not using compact navigation, the navigation bar has two looks. With one extension loaded, it shows the name of the add-in.

![Shows the expanded navigation bar when one module extension is loaded in Outlook.](../images/outlook-module-navigationbar-one.png)

When more than one add-in is loaded, it shows the word **Add-ins**. Clicking either will open the extension's user interface.

![Shows the expanded navigation bar when more than on module extension is loaded in Outlook.](../images/outlook-module-navigationbar-more.png)

When you click on an extension, Outlook replaces the built-in module with your custom module so that your users can interact with the add-in. You can use some of the features of the Outlook JavaScript API in your add-in. APIs that logically assume a specific Outlook item, such as a message or appointment, don't work in module extensions. The module can also include function commands in the Outlook ribbon that interact with the add-in's page. To facilitate this, your function commands call the [Office.onReady or Office.initialize method](../develop/initialize-add-in.md) and the [Event.completed](/javascript/api/office/office.addincommands.event#office-office-addincommands-event-completed-member(1)) method. To walk through how a module extension Outlook add-in is configured, see the [Outlook module extensions billable hours sample](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension).

The following screenshot shows an add-in that is integrated in the Outlook navigation bar and has ribbon commands that will update the page of the add-in.

![Shows the user interface of a module extension.](../images/outlook-module-extension.png)

## Example

The following is a section of a manifest file that defines a module extension.

```xml
<!-- Add Outlook module extension point -->
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                  xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
                    xsi:type="VersionOverridesV1_1">

    <!-- Begin override of existing elements -->
    <Description resid="residVersionOverrideDesc" />

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <!-- End override of existing elements -->

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <!-- Set the URL of the file that contains the
                JavaScript function that controls the extension -->
          <FunctionFile resid="residFunctionFileUrl" />

          <!--New Extension Point - Module for a ModuleApp -->
          <ExtensionPoint xsi:type="Module">
            <SourceLocation resid="residExtensionPointUrl" />
            <Label resid="residExtensionPointLabel" />

            <CommandSurface>
              <CustomTab id="idTab">
                <Group id="idGroup">
                  <Label resid="residGroupLabel" />

                  <Control xsi:type="Button" id="group.changeToAssociate">
                    <Label resid="residChangeToAssociateLabel" />
                    <Supertip>
                      <Title resid="residChangeToAssociateLabel" />
                      <Description resid="residChangeToAssociateDesc" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="residAssociateIcon16" />
                      <bt:Image size="32" resid="residAssociateIcon32" />
                      <bt:Image size="80" resid="residAssociateIcon80" />
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>changeToAssociateRate</FunctionName>
                    </Action>
                  </Control>
                  
              </Group>
                <Label resid="residCustomTabLabel" />
              </CustomTab>
            </CommandSurface>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="residAddinIcon16" 
                  DefaultValue="https://localhost:8080/Executive-16.png" />
        <bt:Image id="residAddinIcon32" 
                  DefaultValue="https://localhost:8080/Executive-32.png" />
        <bt:Image id="residAddinIcon80" 
                  DefaultValue="https://localhost:8080/Executive-80.png" />
      
        <bt:Image id="residAssociateIcon16" 
                  DefaultValue="https://localhost:8080/Associate-16.png" />
        <bt:Image id="residAssociateIcon32" 
                  DefaultValue="https://localhost:8080/Associate-32.png" />
        <bt:Image id="residAssociateIcon80" 
                  DefaultValue="https://localhost:8080/Associate-80.png" />
      </bt:Images>

      <bt:Urls>
        <bt:Url id="residFunctionFileUrl" 
                DefaultValue="https://localhost:8080/" />
        <bt:Url id="residExtensionPointUrl" 
                DefaultValue="https://localhost:8080/" />
      </bt:Urls>

      <!--Short strings must be less than 30 characters long -->
      <bt:ShortStrings>
        <bt:String id="residExtensionPointLabel" 
                    DefaultValue="Billable Hours" />
        <bt:String id="residGroupLabel" 
                    DefaultValue="Change billing rate" />
        <bt:String id="residCustomTabLabel" 
                    DefaultValue="Billable hours" />

        <bt:String id="residChangeToAssociateLabel" 
                    DefaultValue="Associate" />
      </bt:ShortStrings>

      <bt:LongStrings>
        <bt:String id="residVersionOverrideDesc" 
                    DefaultValue="Version override description" />

        <bt:String id="residChangeToAssociateDesc" 
                    DefaultValue="Change to the associate billing rate: $127/hr" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

## See also

- [Outlook add-in manifests](manifests.md)
- [Add-in commands for Outlook](add-in-commands-for-outlook.md)
- [Outlook module extensions Billable hours sample](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension)
