---
title: Use Smart Alerts and the OnMessageSend event in your Outlook add-in (preview)
description: Learn how to handle the send message event in your Outlook add-in using event-based activation.
ms.topic: article
ms.date: 12/22/2021
ms.localizationpriority: medium
---

# Use Smart Alerts and the OnMessageSend event in your Outlook add-in (preview)

The `OnMessageSend` event takes advantage of Smart Alerts which allow you to run logic after a user selects **Send** in their Outlook message. Your event handler allows you to give your users the opportunity to improve their emails before they're sent. The `OnAppointmentSend` event is similar but applies to an appointment.

By the end of this walkthrough, you'll have an add-in that runs whenever a message is being sent and checks if the user forgot to add a document or picture they mentioned in their email.

> [!IMPORTANT]
> The `OnMessageSend` and `OnAppointmentSend` events are only available in preview with a Microsoft 365 subscription in Outlook on Windows. For more details, see [How to preview](autolaunch.md#how-to-preview). Preview events shouldn't be used in production add-ins.

## Prerequisites

The `OnMessageSend` event is available through the event-based activation feature. To understand about configuring your add-in to use this feature, available events, how to preview this event, debugging, feature limitations, and more, refer to [Configure your Outlook add-in for event-based activation](autolaunch.md).

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.

## Configure the manifest

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
                <Control xsi:type="Button" id="ActionButton">
                  <Label resid="ActionButton.Label"/>
                  <Supertip>
                    <Title resid="ActionButton.Label"/>
                    <Description resid="ActionButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
        <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
        <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
>
> - For **SendMode** options available with the `OnMessageSend` event, refer to [Available SendMode options](../reference/manifest/launchevent.md#available-sendmode-options-preview).
> - To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).

## Implement event handling

You have to implement handling for your selected event.

In this scenario, you'll add handling for sending a message. Your add-in will check for certain keywords in the message. If any of those keywords are found, it will then check if there are any attachments. If there are no attachments, your add-in will recommend to the user to add the possibly missing attachment.

1. From the same quick start project, open the file **./src/commands/commands.js** in your code editor.

1. After the `action` function, insert the following JavaScript functions.

    ```js
    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { "asyncContext": event },
        function (asyncResult) {
          var event = asyncResult.asyncContext;
          var body = "";
          if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
            body = asyncResult.value;
          }
        
          var arrayOfTerms = ["send", "picture", "document", "attachment"];
          for (var index = 0; index < arrayOfTerms.length; index++) {
            var term = arrayOfTerms[index].trim();
            const regex = RegExp(term, 'i');
            if (regex.test(body)) {
              matches.push(term);
            }
          }
        
          if (matches.length > 0) {
            // Let's verify if there's an attachment!
            Office.context.mailbox.item.getAttachmentsAsync(
              { "asyncContext": event },
              function(result){
                var event = asyncResult.asyncContext;
                if (result.value.length <= 0) {
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                } else {
                  for (var i=0;i<result.value.length;i++) {
                    if(result.value[i].isInline == false) {
                      event.completed({ allowEvent: true });
                      return;
                    }
                  }
                    
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                }
              });
            } else {
              event.completed({ allowEvent: true });
            }
          }
        );
    }
    ```

1. Add the following JavaScript code at the end of the file.

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. Save your changes.

> [!IMPORTANT]
> Windows: At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.

## Try it out

1. Run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.

1. In Outlook on Windows, create a new message and set the subject. In the body, add text like "Hey, check out this picture of my dog!".
1. Send the message. A dialog should pop up with a recommendation for you to add an attachment.
1. Add an attachment then send the message again. There should be no alert this time.

> [!NOTE]
> If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*. Make sure you have a network connection. If the problem continues, please try again later.", you may need to enable a loopback exemption.
>
> 1. Close Outlook.
> 1. Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.
> 1. If you're using `https://localhost` (the default version in the manifest), run the following command.
>
>    ```command&nbsp;line
>    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E5
>    ```
>
> 1. If you're using `http://localhost`, run the following command.
>
>    ```command&nbsp;line
>    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E5
>    ```
>
> 1. Restart Outlook.

## See also

- [Outlook add-in manifests](manifests.md)
- [Configure your Outlook add-in for event-based activation](autolaunch.md)
- [How to debug event-based add-ins](debug-autolaunch.md)
- [AppSource listing options for your event-based Outlook add-in](autolaunch-store-options.md)
