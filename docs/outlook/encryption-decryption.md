---
title: Create an encryption Outlook add-in
description: Learn how to develop an Outlook add-in that encrypts and decrypts messages.
ms.date: 06/30/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Create an encryption Outlook add-in

Implement custom encryption and decryption functionality in an Outlook add-in to secure email communications. The `OnMessageDecrypt` event lets your add-in automatically identify encrypted messages and handle decryption, content display, and error notifications.

## Overview of the encryption and decryption workflows

> [!TIP]
>
> - The encryption and decryption workflows implement the event-based activation feature. If you aren't familiar with event-based activation in Outlook add-ins, we recommend that you first learn about the feature and its implementation. To learn more, see [Activate add-ins with events](../develop/event-based-activation.md).
> - The minimum requirement set and supported platforms may vary for each API recommended in this section. We recommend verifying any requirements against [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support) and supplementing it with documentation for the specific API.

The following table provides an overview of the encryption and decryption workflows of an Outlook add-in. It also identifies whether a step requires a custom solution or is supported by the Office JavaScript (Office.js) API library.

| Step | Implementation |
| ---- | -------------- |
| User composes a message and uses your add-in to apply encryption rules | You must implement your own encryption protocol so that the add-in can secure the contents of the message and its attachments. |
| User sends the message | Implement a handler for the [OnMessageSend](onmessagesend-onappointmentsend-events.md) event so that your add-in can automatically run your encryption protocol when the user selects **Send**.<br><br>To identify a message that was encrypted using your add-in during the decryption process, use the [internet headers APIs](/javascript/api/outlook/office.internetheaders) to add a header to a message. The header key must match the value specified in the `HeaderName` attribute of the [\<LaunchEvent\>](/javascript/api/manifest/launchevent) element for the [OnMessageDecrypt](../develop/event-based-activation.md#outlook-events) event in the add-in's manifest. For more information, see [Implement decryption using event-based activation](#implement-decryption-using-event-based-activation). |
| Recipient receives the encrypted message and opens it | If the recipient has the same add-in that was used to encrypt the message installed in Outlook, the add-in checks whether the header key included in the message matches the value specified for the `OnMessageDecrypt` event in the manifest. This operation is automatically done by an add-in that handles the `OnMessageDecrypt` event, so that you don't have to manually implement the check. If the headers match, the `OnMessageDecrypt` event occurs and its handler runs. For more information, see [Implement decryption using event-based activation](#implement-decryption-using-event-based-activation). |
| Add-in decrypts the message | You must implement your own decryption protocol in the `OnMessageDecrypt` event handler. While your add-in decrypts the message and its attachments, a notification is shown to the user to alert them that their message is being processed by the add-in. This notification is automatically shown by an add-in that handles the `OnMessageDecrypt` event, so that you don't have to manually create one. |
| Recipient views the decrypted message and its attachments, if any | Once the decryption operation is complete, a notification is automatically shown to the user to alert them that the add-in has finished processing the message. In your `OnMessageDecrypt` handler, call the [event.completed](/javascript/api/outlook/office.mailboxevent#outlook-office-mailboxevent-completed-member(1)) method and pass it a [MessageDecryptEventCompletedOptions](/javascript/api/outlook/office.messagedecrypteventcompletedoptions) object. With the `MessageDecryptEventCompletedOptions` object, you can specify whether to display the decrypted content to the recipient. For more information, see [Implement event handling](#implement-event-handling). |

## Try out a completed add-in

To immediately see a completed encryption add-in in action, try out the [Encrypt and decrypt messages in Outlook sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-decrypt-messages).

## Implement decryption using event-based activation

You must implement your own encryption and decryption protocols. The add-in must also be configured to handle the `OnMessageDecrypt` event to conveniently determine when your add-in can decrypt a message and display the decrypted contents. To implement the `OnMessageDecrypt` event, you must:

1. [Configure the add-in's manifest](#configure-the-manifest).
1. [Implement event handling](#implement-event-handling).

### Supported environments

The `OnMessageDecrypt` event is supported on the Message Read surface. Support varies by client and Exchange environment, as shown in the following table.

| Client | Exchange Online | Exchange Subscription Edition (SE) | Exchange Server 2019 | Exchange Server 2016 |
| ------ | --------------- | ---------------------------------- | -------------------- | -------------------- |
| **Web browser** | Supported | Not available | Not available | Not available |
| **Windows (new)** | Supported | Not available | Not available | Not available |
| **Windows (classic)**<br>Version 2602 (Build 19725.20126) and later | Supported | Not available | Not available | Not available |
| **Mac** | Not available | Not available | Not available | Not available |
| **Android** | Not available | Not available | Not available | Not available |
| **iOS** | Not available | Not available | Not available | Not available |

### Configure the manifest

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

> [!NOTE]
> The `OnMessageDecrypt` event and `"extensions.autoRunEvents.events.options.headerName"` property are in preview with the unified manifest. Don't use the decryption feature with the unified manifest in a production add-in.

In your add-in's **manifest.json** file, you must configure the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array) array and add the [`"extensions.autoRunEvents"`](/microsoft-365/extensibility/schema/element-extensions#autorunevents) array to enable event-based activation in your add-in.

1. Add the following object to the `"extensions.runtimes"` array. Note the following about this markup.

     - The `"id"` of the runtime is set to the descriptive name `"autorun_runtime"`.
     - The `"code"` property has a child `"page"` property that is set to an HTML file and a child `"script"` property that is set to a JavaScript file. Office uses one of these values depending on the platform.
         - Outlook on the web and the [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) execute the handler in a browser runtime, which loads an HTML file. That file, in turn, contains a `<script>` tag that loads the JavaScript file.
         - Classic Outlook on Windows executes the event handler in a JavaScript-only runtime, which loads a JavaScript file directly.
         For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).
     - The `"lifetime"` property is set to `"short"`, which means that the runtime starts up when the event is triggered and shuts down when the handler completes.
     - [Actions](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) map JavaScript handlers to the `OnMessageSend` and `OnMessageDecrypt` events.

    ```json
    "runtimes": [
        {
            "requirements": {
                "capabilities": [
                    {
                        "name": "Mailbox",
                        "minVersion": "1.16"
                    }
                ]
            },
            "id": "autorun_runtime",
            "type": "general",
            "code": {
                "page": "https://localhost:3000/launchevents.html",
                "script": "https://localhost:3000/launchevents.js"
            },
            "lifetime": "short",
            "actions": [
                {
                    "id": "onMessageSendHandler",
                    "type": "executeFunction"
                },
                {
                    "id": "onMessageDecryptHandler",
                    "type": "executeFunction"
                }
            ]
        }
    ],
    ```

1. Add the following `"autoRunEvents"` array as a property of the object in the `"extensions"` array. Note the following about this markup.

   - An event object is created for each event that the add-in handles. In this sample, one event object is created for `OnMessageSend` and another for `OnMessageDecrypt`. Both events use their unified manifest event name, `"messageSending"` and `"messageDecrypt"`, as described in the [supported events table](../develop/event-based-activation.md#supported-events).
   - To ensure that the appropriate handler runs when an event occurs, the function name provided in `"actionId"` must match the name used in the `"id"` property of the applicable object in the `"runtimes.actions"` array from an earlier step.
   - The ["options"](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events-options) property provides additional configuration for the `OnMessageSend` and `OnMessageDecrypt` events.
       - For `OnMessageSend`, the ["sendMode"](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events-options#sendmode) option specifies whether a user is able to send their message if it doesn't meet an add-in's conditions. In this sample, the `"softBlock"` option is specified. To learn more about send mode options, see the "Available send mode options" section of [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](onmessagesend-onappointmentsend-events.md#available-send-mode-options).
       - For `OnMessageDecrypt`, the ["headerName"](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events-options#headername) option specifies the internet header name used to identify whether a message was encrypted by the add-in. The same header is added to a message that's encrypted by the add-in.

    ```json
    "autoRunEvents": [
        {
            "events": [
              {
                  "type": "messageSending",
                  "actionId": "onMessageSendHandler",
                  "options": {
                      "sendMode": "softBlock"
                  }
              },
              {
                  "type": "messageDecrypt",
                  "actionId": "onMessageDecryptHandler",
                  "options": {
                      "headerName": "contoso-encrypted"
                  }
              }
            ]
        }
    ]
    ```

# [Add-in only manifest](#tab/xmlmanifest)

To activate your add-in on the `OnMessageDecrypt` event, you must configure the [\<VersionOverridesV1_1\>](/javascript/api/manifest/versionoverrides-1-1-mail) node of your add-in's **manifest.xml** file as follows.

- To run an event-based add-in in Outlook on the web and the new Outlook on Windows, you must specify the HTML file that contains your event-handling code in the `resid` attribute of the [\<Runtime\>](/javascript/api/manifest/runtime) element. To run in classic Outlook on Windows, you must also specify the JavaScript file that contains the event handler in the [\<Override\>](/javascript/api/manifest/override) child element of the `<Runtime>` element.

  > [!TIP]
  > To learn more about runtimes, see [Runtimes in Office Add-ins](../testing/runtimes.md).
- Specify the `OnMessageDecrypt` event in the `Type` attribute of a [\<LaunchEvent\>](/javascript/api/manifest/launchevent) element. You must assign the function name of the event handler to the `FunctionName` attribute of the element. To facilitate checking whether the message was encrypted by the add-in, a header key must be specified in the `HeaderName` attribute. The same header is added to a message that's encrypted by the add-in.

The following is an example of a `<VersionOverrides>` node that implements the `OnMessageDecrypt` event.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.15">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <Runtimes>
            <!-- References the HTML file that links to the event handler. This is used by Outlook on the web and the new Outlook on Windows. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- References the JavaScript file that contains the event handler. This is used by classic Outlook on Windows. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="WebViewRuntime.Url"/>
          <!-- Implements event-based activation. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock"/>
              <LaunchEvent Type="OnMessageDecrypt" FunctionName="onMessageDecryptHandler" HeaderName="contoso-encrypted"/>
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
        </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      ...
      <bt:Urls>
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://www.contoso.com/launchevent.html"/>
        <bt:Url id="JSRuntime.Url" DefaultValue="https://www.contoso.com/launchevent.js"/>
      </bt:Urls>
      ...
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

---

### Implement event handling

The `OnMessageDecrypt` event handler is used to run the decryption operation and determine whether to display the decrypted contents of a message.

- To ensure your handler runs when the `OnMessageDecrypt` event occurs, call `Office.actions.associate` in the JavaScript file where the handler is implemented. This maps the handler name specified in the `FunctionName` attribute of the `<LaunchEvent>` element in the manifest to its JavaScript counterpart.
- Once the decryption operation finishes, you must call `event.completed` to signal to the client that your add-in has completed processing the `OnMessageDecrypt` event. To display the decrypted contents of a message and its attachments, pass a [MessageDecryptEventCompletedOptions](/javascript/api/outlook/office.messagedecrypteventcompletedoptions) object to the `event.completed` call and set its [allowEvent](/javascript/api/outlook/office.messagedecrypteventcompletedoptions#outlook-office-messagedecrypteventcompletedoptions-allowevent-member) property to `true`. Then, specify the decrypted contents of the message in the object's [emailBody](/javascript/api/outlook/office.messagedecrypteventcompletedoptions#outlook-office-messagedecrypteventcompletedoptions-emailbody-member) and [attachments](/javascript/api/outlook/office.messagedecrypteventcompletedoptions#outlook-office-messagedecrypteventcompletedoptions-attachments-member) properties. You can also specify any data that your add-in may need for processing in the [contextData](/javascript/api/outlook/office.messagedecrypteventcompletedoptions#outlook-office-messagedecrypteventcompletedoptions-contextdata-member) property. For example, you can store custom internet headers to decrypt messages in reply and forward scenarios.

> [!NOTE]
> Be mindful of the following when creating an event-based add-in for classic Outlook on Windows.
>
> - Imports aren't currently supported in the JavaScript file containing the event handler.
> - When the JavaScript function specified in the manifest to handle an event runs, code in `Office.onReady()` and `Office.initialize` isn't run. We recommend adding any startup logic needed by the event handler, such as checking the user's Outlook version, to the event handler instead.

The following is an example of an `OnMessageDecrypt` event handler.

```javascript
function onMessageDecryptHandler(event) {
    // Your code to decrypt the contents of a message would appear here.
    ...

    // Use the results from your decryption process to display the decrypted contents of the message body and attachments.
    const decryptedBodyContent = "<p>Please find attached the recent report and its supporting documentation.</p>";
    const decryptedBody = {
        coercionType: Office.CoercionType.Html,
        content: decryptedBodyContent
    };

    // Decrypted content and properties of a file attachment.
    const decryptedPdfFile = "JVBERi0xLjQKJeLjz9MKNCAwIG9i...";
    const pdfFileName = "Fabrikam_Report_202509";

    // Decrypted properties of a cloud attachment.
    const cloudFilePath = "https://contosostorage.com/reports/weekly_forecast.xlsx";
    const cloudFileName = "weekly_forecast.xlsx";

    // Decrypted content and properties of an inline image.
    const decryptedImageFile = "iVBORw0KGgoAAAANSUhEUgAA...";
    const imageFileName = "banner.png";
    const imageContentId = "image001.png@01DC1DD9.1A4AA300";

    const decryptedAttachments = [
        {
            attachmentType: Office.MailboxEnums.AttachmentType.File,
            content: decryptedPdfFile,
            isInline: false,
            name: pdfFileName
        },
        {
            attachmentType: Office.MailboxEnums.AttachmentType.Cloud,
            isInline: false,
            name: cloudFileName,
            path: cloudFilePath
        },
        {
            attachmentType: Office.MailboxEnums.AttachmentType.File,
            content: decryptedImageFile,
            contentId: imageContentId,
            isInline: true,
            name: imageFileName
        }
    ];

    event.completed({
        allowEvent: true,
        emailBody: decryptedBody,
        attachments: decryptedAttachments,
        contextData: { messageType: "ReplyFromDecryptedMessage" }
    });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageDecryptHandler", onMessageDecryptHandler);
```

> [!TIP]
> When images are added to a message as inline attachments, they're automatically assigned a content ID. In the body of a message, the content ID of an inline attachment is specified in the `src` attribute of the `<img>` element similar to the following example.
>
> `<img width=96 height=96 id="Picture_1" src="cid:image001.png@01DC1E6F.FC7C7410">`
>
> To easily identify and provide these inline attachments during decryption, we recommend saving the content IDs of inline attachments to the message header during encryption. Call [Office.context.mailbox.item.getAttachmentsAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getattachmentsasync-member(1)) to get the [content ID](/javascript/api/outlook/office.attachmentdetailscompose#outlook-office-attachmentdetailscompose-contentid-member) of an inline attachment. Then, call [Office.context.mailbox.item.internetHeaders.setAsync](/javascript/api/outlook/office.internetheaders#outlook-office-internetheaders-setasync-member(1)) to save the ID to the header of the message.

#### Decrypt Outlook item attachments (preview)

Support for decrypting Outlook item attachments (`Office.MailboxEnums.AttachmentType.Item`), particularly email attachments, is available for preview in Outlook on the web and on Windows (new and classic). To preview this feature in classic Outlook on Windows, you must install Version 2606 (Build 20114.15110) or later. Then, join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/kb/microsoft-365-insider-kb/join-the-microsoft-365-insider-program-on-windows/4401748) and select the **Beta Channel** option to access Office beta builds. To test this feature using the sample code in this article, update the `onMessageDecryptHandler` function with the following code.

```javascript
    // Decrypted content and properties of an email attachment.
    const decryptedEmailFile = "VGhpcyBpcyBhIHRleHQgZmlsZS4=...";
    const emailFileName = "Fabrikam_Report_202508.eml";

    const decryptedAttachments = [
        ...
        {
            attachmentType: Office.MailboxEnums.AttachmentType.Item,
            content: decryptedEmailFile,
            name: emailFileName
        }
    ];
    ...
```

#### Customize error messages for the decryption operation (preview)

Custom error messages for failed decryption operations are available for preview in Outlook on the web and on Windows (new and classic). To preview this feature in classic Outlook on Windows, you must install Version 2606 (Build 20114.15110) or later. Then, join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/kb/microsoft-365-insider-kb/join-the-microsoft-365-insider-program-on-windows/4401748) and select the **Beta Channel** option to access Office beta builds.

If the decryption operation fails, the `allowEvent` property of the `event.completed` call is set to `false`, and Outlook shows the following default notification to the user: "\<Add-in name\> failed to process your message." To specify a custom error message, set the [errorMessage](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-messagedecrypteventcompletedoptions-errormessage-member) property of your add-in's `event.completed` call. Your custom message is prefixed with "Error from \<add-in name\>:". If your custom message can't be shown, the default notification is shown instead.

The following code sample shows how to specify a custom error message for your decryption add-in.

```javascript
event.completed({
    allowEvent: false,
    errorMessage: "This message couldn't be decrypted. Contact the Contoso IT team for further assistance."
});
```

#### Manage distribution of decrypted content (preview)

To help prevent unauthorized distribution of decrypted content, access control options are available for preview in Outlook on the web and on Windows (new and classic). To preview this feature in classic Outlook on Windows, you must install Version 2606 (Build 20114.15110) or later. Then, join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/kb/microsoft-365-insider-kb/join-the-microsoft-365-insider-program-on-windows/4401748) and select the **Beta Channel** option to access Office beta builds.

To limit printing, copying, or saving of decrypted content, include the [accessControls](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-messagedecrypteventcompletedoptions-accesscontrols-member) property of the `event.completed` call. Then, set the [allowPrint](/javascript/api/outlook/office.accesscontrols?view=outlook-js-preview&preserve-view=true#outlook-office-accesscontrols-allowprint-member), [allowCopyPaste](/javascript/api/outlook/office.accesscontrols?view=outlook-js-preview&preserve-view=true#outlook-office-accesscontrols-allowcopypaste-member), and [allowSave](/javascript/api/outlook/office.accesscontrols?view=outlook-js-preview&preserve-view=true#outlook-office-accesscontrols-allowsave-member) properties to `false`. If the `accessControls` property isn't specified, access controls default to `true`.

To test this feature using the sample code in this article, update the `event.completed` call of the `onMessageDecryptHandler` function with the following code.

```javascript
    event.completed({
        allowEvent: true,
        emailBody: decryptedBody,
        attachments: decryptedAttachments,
        contextData: { messageType: "ReplyFromDecryptedMessage" },
        accessControls: {
            allowPrint: false,
            allowCopyPaste: false,
            allowSave: false
        }
    });
```

> [!NOTE]
>
> - In Outlook on the web, setting the `allowCopyPaste` property to `false` also prevents users from capturing their screen in the form of screenshots or recordings. The screen capture policy remains in effect until the user reloads the Outlook browser tab.
> - In Outlook on the web and the new Outlook on Windows, setting the `allowCopyPaste` property to `true` allows the user to copy content by selecting **Copy** from the context menu or pressing <kbd>Ctrl</kbd>+<kbd>C</kbd>. However, if another access control is set to `false`, the context menu becomes unavailable. The user must use <kbd>Ctrl</kbd>+<kbd>C</kbd> instead.

## Behavior and limitations

- Be aware of the behaviors and limitations of event-based add-ins. To learn more, see [Activate add-ins with events](../develop/event-based-activation.md#behavior-and-limitations).
- Since each add-in uses its own encryption protocol, a message can only be decrypted by the same add-in that encrypted it. When a user doesn't have the required add-in installed to decrypt a message, a notification alerts them that the message is encrypted. To guide the user through the decryption process, customize a placeholder message for the body of the encrypted message. The placeholder message can include information on how to install your add-in. To set the message body during the encryption process, call [Office.context.mailbox.item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(2)).

    :::image type="content" source="../images/outlook-encryption-placeholder-message.png" alt-text="A sample placeholder message of an encrypted message.":::

- To ensure data security and confidentiality, decrypted content isn't stored on the Outlook client. The contents of an encrypted message are decrypted every time a user opens it.
- An encrypted message must first be decrypted before a user can reply or forward it. A user can't reply or forward an encrypted message while it's being decrypted.
- If a user navigates to another mail item while an encrypted message is being decrypted, the decryption process stops running. The user must select or open the message again to activate the decryption process.
- When replying to or forwarding encrypted messages, drafts are saved unencrypted in the **Drafts** folder.
- The `attachments` property of the `event.completed` method doesn't support attachments of type `Office.MailboxEnums.AttachmentType.Item`, except for preview in Outlook on the web and on Windows (new and classic). To learn more, see [Decrypt Outlook item attachments (preview)](#decrypt-outlook-item-attachments-preview).
- Custom encryption add-ins can't encrypt messages that are already protected by DRM or S/MIME.
- In Outlook on the web and the new Outlook on Windows, when encrypted messages are [grouped by conversation](https://support.microsoft.com/outlook/mail/view-email-messages-by-conversation-in-outlook), only the currently selected message from the conversation thread is decrypted. The other messages in the conversation thread remain encrypted until they're selected.
- In Outlook on the web and the new Outlook on Windows, users can only download a decrypted message in the EML format. The option to download in the MSG format is unavailable.

### Decryption notifications

Add-ins that handle the `OnMessageDecrypt` event automatically display notifications in certain decryption scenarios as described in the following table.

| Notification | Scenario |
| ------------ | -------- |
| \<Add-in name\> is unavailable and can't process your message at this time. | Applies to classic Outlook on Windows only. This notification is shown when the add-in fails to load because an error prevented the add-in from loading or the user's client or machine is offline. |
| \<Add-in name\> failed to process your message. | An error was encountered while the add-in was decrypting the message. To retry the decryption operation, the recipient must switch to another message, then open the encrypted message again to invoke the `OnMessageDecrypt` event. |
| \<Add-in name\> add-in is decrypting your message. | The add-in is handling the `OnMessageDecrypt` event to decrypt the message. |
| This message is encrypted by \<add-in name\> add-in. | This notification is shown to recipients who don't have the necessary encryption add-in installed. To provide guidance on how to decrypt the message, include a placeholder message in the body of the encrypted message. For more information, see [Behavior and limitations](#behavior-and-limitations). |
| \<Add-in name\> add-in has decrypted your message. | The add-in successfully decrypted the contents of the message. The user can now view the message and its attachments. |
| \<Add-in name\> is taking longer than expected to process your message. | The add-in has been running for more than five seconds, but less than five minutes. |
| \<Add-in name\> timed out. To retry, select another email and then return to this message. | The add-in times out after running for five minutes. To retry the decryption operation, the recipient must switch to another message, then open the encrypted message again to invoke the `OnMessageDecrypt` event. |
| \<Add-in name\> timed out. (preview) | The add-in times out after running for five minutes. This notification includes a **Retry** action so the recipient can retry the decryption operation without switching to another message. This retry feature is available for preview in Outlook on the web and on Windows (new and classic). To preview this feature in classic Outlook on Windows, you must install Version 2606 (Build 20114.15110) or later. Then, join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/kb/microsoft-365-insider-kb/join-the-microsoft-365-insider-program-on-windows/4401748) and select the **Beta Channel** option to access Office beta builds. |
| \<Add-in name\> can't process this message because it's protected by a built-in security feature. | The add-in tries to process a message that's already protected by DRM or S/MIME. |
| Custom error message (preview) | An error was encountered while the add-in was decrypting the message. To retry the decryption operation, the recipient must switch to another message, then open the encrypted message again to invoke the `OnMessageDecrypt` event. For guidance on how to customize an error message for the decryption operation, see [Customize error messages for the decryption operation (preview)](#customize-error-messages-for-the-decryption-operation-preview). |

## See also

- [Sample: Encrypt and decrypt messages in Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-decrypt-messages)
- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
- [Activate add-ins with events](../develop/event-based-activation.md)
- [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md)
- [Get and set internet headers on a message in an Outlook add-in](internet-headers.md)
- [Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)
