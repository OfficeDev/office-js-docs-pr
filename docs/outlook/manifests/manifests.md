
# Outlook add-in manifests

An Outlook add-in consists of two components: the XML add-in manifest, and a web page, supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. Currently there are three versions of the manifest schema, including  **VersionOverrides**. We recommend that you use manifest schema version 1.1 and  **VersionOverrides** 1.0 to build your add-in. The following is an example.

 >**Note**  All URL values in the following sample begin with "YOUR_WEB_SERVER". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.




```XML
<?xml version="1.0" encoding="UTF-8"?>
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
  <IconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
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
        <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
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
          
          <!-- Custom pane, only applies to read form -->
          <ExtensionPoint xsi:type="CustomPane">
            <RequestedHeight>100</RequestedHeight> 
            <SourceLocation resid="customPaneUrl"/>
            <Rule xsi:type="RuleCollection" Mode="Or">
              <Rule xsi:type="ItemIs" ItemType="Message"/>
              <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee"/>
            </Rule>
          </ExtensionPoint>
          
          <!-- Message compose form -->
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgComposeDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgComposeFunctionButton">
                  <Label resid="funcComposeButtonLabel" />
                  <Tooltip resid="funcComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="funcComposeSuperTipTitle" />
                    <Description resid="funcComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>addDefaultMsgToBody</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgComposeMenuButton">
                  <Label resid="menuComposeButtonLabel" />
                  <Tooltip resid="menuComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="menuComposeSuperTipTitle" />
                    <Description resid="menuComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgComposeMenuItem1">
                      <Label resid="menuItem1ComposeLabel" />
                      <Tooltip resid="menuItem1ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem1ComposeLabel" />
                        <Description resid="menuItem1ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg1ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgComposeMenuItem2">
                      <Label resid="menuItem2ComposeLabel" />
                      <Tooltip resid="menuItem2ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem2ComposeLabel" />
                        <Description resid="menuItem2ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg2ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgComposeMenuItem3">
                      <Label resid="menuItem3ComposeLabel" />
                      <Tooltip resid="menuItem3ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem3ComposeLabel" />
                        <Description resid="menuItem3ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg3ToBody</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                  <Label resid="paneComposeButtonLabel" />
                  <Tooltip resid="paneComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="paneComposeSuperTipTitle" />
                    <Description resid="paneComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="composeTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          
          <!-- Appointment compose form -->
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="apptComposeFunctionButton">
                  <Label resid="funcComposeButtonLabel" />
                  <Tooltip resid="funcComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="funcComposeSuperTipTitle" />
                    <Description resid="funcComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>addDefaultMsgToBody</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="apptComposeMenuButton">
                  <Label resid="menuComposeButtonLabel" />
                  <Tooltip resid="menuComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="menuComposeSuperTipTitle" />
                    <Description resid="menuComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="apptComposeMenuItem1">
                      <Label resid="menuItem1ComposeLabel" />
                      <Tooltip resid="menuItem1ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem1ComposeLabel" />
                        <Description resid="menuItem1ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg1ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="apptComposeMenuItem2">
                      <Label resid="menuItem2ComposeLabel" />
                      <Tooltip resid="menuItem2ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem2ComposeLabel" />
                        <Description resid="menuItem2ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg2ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="apptComposeMenuItem3">
                      <Label resid="menuItem3ComposeLabel" />
                      <Tooltip resid="menuItem3ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem3ComposeLabel" />
                        <Description resid="menuItem3ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg3ToBody</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                  <Label resid="paneComposeButtonLabel" />
                  <Tooltip resid="paneComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="paneComposeSuperTipTitle" />
                    <Description resid="paneComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="composeTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          
          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Tooltip resid="funcReadButtonTooltip" />
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
                  <Tooltip resid="menuReadButtonTooltip" />
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
                      <Tooltip resid="menuItem1ReadTip" />
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
                      <Tooltip resid="menuItem2ReadTip" />
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
                      <Tooltip resid="menuItem3ReadTip" />
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
                  <Tooltip resid="paneReadButtonTooltip" />
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
          
          <!-- Appointment read form -->
          <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptReadDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="apptReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Tooltip resid="funcReadButtonTooltip" />
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
                <Control xsi:type="Menu" id="apptReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Tooltip resid="menuReadButtonTooltip" />
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
                    <Item id="apptReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Tooltip resid="menuItem1ReadTip" />
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
                    <Item id="apptReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Tooltip resid="menuItem2ReadTip" />
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
                    <Item id="apptReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Tooltip resid="menuItem3ReadTip" />
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
                <Control xsi:type="Button" id="apptReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Tooltip resid="paneReadButtonTooltip" />
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
        <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png"/>
        <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png"/>
        <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png"/>
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="YOUR_WEB_SERVER/images/red-16.png"/>
        <bt:Image id="red-icon-32" DefaultValue="YOUR_WEB_SERVER/images/red-32.png"/>
        <bt:Image id="red-icon-80" DefaultValue="YOUR_WEB_SERVER/images/red-80.png"/>
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="YOUR_WEB_SERVER/images/green-16.png"/>
        <bt:Image id="green-icon-32" DefaultValue="YOUR_WEB_SERVER/images/green-32.png"/>
        <bt:Image id="green-icon-80" DefaultValue="YOUR_WEB_SERVER/images/green-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
        <bt:Url id="readTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <bt:Url id="composeTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppCompose/TaskPane/TaskPane.html"/>
        <bt:Url id="customPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/CustomPane/CustomPane.html"/>"
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo"/>
        <!-- Compose mode -->
        <bt:String id="funcComposeButtonLabel" DefaultValue="Insert default message"/>
        <bt:String id="menuComposeButtonLabel" DefaultValue="Insert message"/>
        <bt:String id="paneComposeButtonLabel" DefaultValue="Insert custom message"/>

        <bt:String id="funcComposeSuperTipTitle" DefaultValue="Inserts the default message"/>
        <bt:String id="menuComposeSuperTipTitle" DefaultValue="Choose a message to insert"/>
        <bt:String id="paneComposeSuperTipTitle" DefaultValue="Enter your own text to insert"/>
        
        <bt:String id="menuItem1ComposeLabel" DefaultValue="Hello World!"/>
        <bt:String id="menuItem2ComposeLabel" DefaultValue="Add-in commands are cool!"/>
        <bt:String id="menuItem3ComposeLabel" DefaultValue="Visit dev.outlook.com"/>

        <!-- Read mode -->
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject"/>
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property"/>
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties"/>

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment"/>
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get"/>
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties"/>
        
        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class"/>
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created"/>
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="groupTooltip" DefaultValue="Add-in Demo of the different command types"/>
        <!-- Compose mode -->
        <bt:String id="funcComposeButtonTooltip" DefaultValue="Inserts text into body of the message or appointment"/>
        <bt:String id="menuComposeButtonTooltip" DefaultValue="Inserts your choice of text into body of the message or appointment"/>
        <bt:String id="paneComposeButtonTooltip" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment"/>
        
        <bt:String id="funcComposeSuperTipDescription" DefaultValue="Inserts text into body of the message or appointment. This is an example of a function button."/>
        <bt:String id="menuComposeSuperTipDescription" DefaultValue="Inserts your choice of text into body of the message or appointment. This is an example of a drop-down menu button."/>
        <bt:String id="paneComposeSuperTipDescription" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment. This is an example of a button that opens a task pane."/>
        
        <bt:String id="menuItem1ComposeTip" DefaultValue="Inserts Hello World! into the body of the message or appointment." />
        <bt:String id="menuItem2ComposeTip" DefaultValue="Inserts Add-in commands are cool! into the body of the message or appointment." />
        <bt:String id="menuItem3ComposeTip" DefaultValue="Inserts Visit dev.outlook.com into the body of the message or appointment." />

        <!-- Read mode -->
        <bt:String id="funcReadButtonTooltip" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar"/>
        <bt:String id="menuReadButtonTooltip" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar"/>
        <bt:String id="paneReadButtonTooltip" DefaultValue="Opens a pane displaying all available properties of the message or appointment"/>
        
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button."/>
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button."/>
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane."/>
        
        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>

```


## Schema versions

Not all Outlook clients support the latest features at once and because some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

The  **VersionOverrides** element in the manifest is an example of this. All elements defined inside **VersionOverrides** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **VersionOverrides** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **VersionOverrides**, Outlook will ignore it and depend on the information in the rest of the manifest. 

This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.

The current versions of the schema are:


|||
|:-----|:-----|
|Version|Description|
|v1.0|Supports version 1.0 of the JavaScript API for Office. For Outlook add-ins, this supports read form. |
|v1.1|Supports version 1.1 of the JavaScript API for Office and  **VersionOverrides**. For Outlook add-ins, this adds support for compose form.|
|**VersionOverrides** 1.0|Supports later versions of the JavaScript API for Office. This supports add-in commands.|
This article will cover the requirements for a 1.1 manifest. Even if your add-in manifest uses the  **VersionOverrides** element, it is still important to include the 1.1 manifest elements to allow your add-in to work with older clients that do not support **VersionOverrides**.


## Root element

The root element for the Outlook add-in manifest is  **OfficeApp**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element:


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest>

</OfficeApp>
```


## Version

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it was approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.


## VersionOverrides

The  **VersionOverrides** element is the location of information for add-in commands. For more information on this element, see [Define add-in commands in your Outlook add-in manifest](../../outlook/manifests/define-add-in-commands.md).


## Localization

Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the  **Resources** element within the **VersionOverrides** element. The following shows how to override an image, a URL, and a string:


```XML
<Resources>
    <bt:Images>
      <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
        <!-- add information for other locales -->

    <bt:Urls>
      <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
        <!-- add information for other locales -->

    <bt:ShortStrings> 
      </bt:String>
      <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
        <bt:Override Locale="ar-sa" Value="<add localized value here>" />
        <!-- add information for other locales -->
    </bt:ShortStrings>

  </Resources>
```

The schema reference contains full information on which elements can be localized.


## Hosts

Outlook add-ins specify the  **Hosts** element like the followiing.


```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

This is separate from the  **Hosts** element inside the **VersionOverrides** element, which is discussed in [Define add-in commands in your Outlook add-in manifest](../../outlook/manifests/define-add-in-commands.md).


## Requirements

The  **Requirements** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](../../outlook/apis.md) for more information on requirement sets.

The  **Requirements** element can also appear in the **VersionOverrides** element, allowing the add-in to specify a different requirement when loaded in clients that support **VersionOverrides**.

The following example uses the  **DefaultMinVersion** attribute of the **Sets** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **Set** element to require the Mailbox requirement set version 1.1.




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

The  **FormSettings** element is used by older Outlook clients, which only support schema 1.1 and not **VersionOverrides**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**.  **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.

These settings are directly related to the activation rules in the  **Rule** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.

For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](http://msdn.microsoft.com/library/7e0cadc3-f613-8eb9-7ef-9032cbb97f92.aspx).


## App domains

The domain of the add-in start page that you specify in the  **SourceLocation** element is the default domain for the add-in. Without using the **AppDomains** and **AppDomain** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **AppDomains** element and include each additional domain in its own **AppDomain** sub-element in the add-in manifest.

The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane:




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


## Permissions

The  **Permissions** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](../../../reference/outlook/Office.context.mailbox.item.md), and does not call [mailbox.makeEwsRequestAsync](../../../reference/outlook/Office.context.mailbox.md) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).


**Four-tier permissions model for mail add-ins**

![4-tier permissions model for mail apps schema v1.1](../../../images/olowa15wecon_Permissions_4Tier.png)
```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>

```


## Activation rules

Activation rules are specified in the  **Rule** element. The **Rule** element can appear as a child of the **OfficeApp** element in 1.1 manifests, and additionally as a child of the **ExtensionPoint** element in **VersionOverrides**. See [Define add-in commands in your Outlook add-in manifest](../../outlook/manifests/define-add-in-commands.md) for details on using this element in **VersionOverrides**.

Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.


- The item type and/or message class
    
- The presence of a specific type of known entity, such as an address or phone number
    
- A regular expression match in the body, subject, or sender email address
    
- The presence of an attachment
    
For details and samples of activation rules, see [Activation rules for Outlook add-ins](../../outlook/manifests/activation-rules.md).


## Next steps: Add-in commands


After defining a basic manifest, [define add-in commands for your add-in](../../outlook/manifests/define-add-in-commands.md). Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).


## Additional resources



- [Outlook add-ins](../../outlook/outlook-add-ins.md)
    
- [Activation rules for Outlook add-ins](../../outlook/manifests/activation-rules.md)
    
- [Localization for Office Add-ins](../../develop/localization.md)
    
- [Create a mail add-in for Outlook that runs on desktops, tablets, and mobile devices (schema v1.1)](http://msdn.microsoft.com/library/8d425fb3-8a7c-429d-87b3-8046e964b153%28Office.15%29.aspx)
    
- [Privacy, permissions, and security for Outlook add-ins](../../outlook/privacy-and-security.md)
    
- [Outlook add-in APIs](../../outlook/apis.md)
    
- [Office Add-ins XML manifest](../../overview/add-in-manifests.md)
    
- [Schema reference for Office Add-ins manifests (v1.1)](http://msdn.microsoft.com/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92%28Office.15%29.aspx)
    
- [Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
    
- [Design guidelines for Office Add-ins](../../design/add-in-design.md)
    
- [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md)
    
- [Use regular expression activation rules to show an Outlook add-in](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Match strings in an Outlook item as well-known entities](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
