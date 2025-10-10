---
title: Create an encryption Outlook add-in (preview)
description: Learn how to develop an Outlook add-in that encrypts and decrypts messages.
ms.date: 10/09/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Create an encryption Outlook add-in (preview)

Implement custom encryption and decryption functionality in an Outlook add-in to secure email communications. The `OnMessageRead` event lets your add-in automatically identify encrypted messages and handle decryption, content display, and error notifications.

> [!NOTE]
> The `OnMessageRead` event and decryption APIs are in preview. Features in preview shouldn't be used in production add-ins as they may change based on feedback we receive. We invite you to try out this feature in test or development environments and welcome feedback on your experience through GitHub (see the "Office Add-ins feedback" section at the end of this page).

## Overview of the encryption and decryption workflows

> [!TIP]
>
> - The encryption and decryption workflows implement the event-based activation feature. If you aren't familiar with event-based activation in Outlook add-ins, we recommend that you first learn about the feature and its implementation. To learn more, see [Activate add-ins with events](../develop/event-based-activation.md).
> - The minimum requirement set and supported platforms may vary for each API recommended in this section. We recommend verifying any requirements against [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support) and supplementing it with documentation for the specific API.

The following table provides an overview of the encryption and decryption workflows of an Outlook add-in. It also identifies whether a step requires a custom solution or is supported by the Office JavaScript (Office.js) API library.

| Step | Implementation |
| ---- | -------------- |
| User composes a message and uses your add-in to apply encryption rules | You must implement your own encryption protocol so that the add-in can secure the contents of the message and its attachments. |
| User sends the message | Implement a handler for the [OnMessageSend](onmessagesend-onappointmentsend-events.md) event so that your add-in can automatically run your encryption protocol when the user selects **Send**.<br><br>To identify a message that was encrypted using your add-in during the decryption process, use the [internet headers APIs](/javascript/api/outlook/office.internetheaders) to add a header to a message. The header key must match the value specified in the `HeaderName` attribute of the [\<LaunchEvent\>](/javascript/api/manifest/launchevent?view=outlook-js-preview&preserve-view=true) element for the [OnMessageRead](../develop/event-based-activation.md#outlook-events) event in the add-in's manifest. For more information, see [Implement decryption using event-based activation](#implement-decryption-using-event-based-activation). |
| Recipient receives the encrypted message and opens it | If the recipient has the same add-in that was used to encrypt the message installed in Outlook, the add-in checks whether the header key included in the message matches the value specified for the `OnMessageRead` event in the manifest. This operation is automatically done by an add-in that handles the `OnMessageRead` event, so that you don't have to manually implement the check. If the headers match, the `OnMessageRead` event occurs and its handler runs. For more information, see [Implement decryption using event-based activation](#implement-decryption-using-event-based-activation). |
| Add-in decrypts the message | You must implement your own decryption protocol in the `OnMessageRead` event handler. While your add-in decrypts the message and its attachments, a notification is shown to the user to alert them that their message is being processed by the add-in. This notification is automatically shown by an add-in that handles the `OnMessageRead` event, so that you don't have to manually create one. |
| Recipient views the decrypted message and its attachments, if any | Once the decryption operation is complete, a notification is automatically shown to the user to alert them that the add-in has finished processing the message. In your `OnMessageRead` handler, call the [event.completed](/javascript/api/outlook/office.mailboxevent?view=outlook-js-preview&preserve-view=true#outlook-office-mailboxevent-completed-member(1)) method and pass it a [MessageDecryptEventCompletedOptions](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true) object. With the `MessageDecryptEventCompletedOptions` object, you can specify whether to display the decrypted content to the recipient. For more information, see [Implement event handling](#implement-event-handling). |

## Implement decryption using event-based activation

You must implement your own encryption and decryption protocols. The add-in must also be configured to handle the `OnMessageRead` event to conveniently determine when your add-in can decrypt a message and display the decrypted contents. To implement the `OnMessageRead` event, you must:

1. [Configure the add-in's manifest](#configure-the-manifest).
1. [Implement event handling](#implement-event-handling).

### Supported environments

The `OnMessageRead` event is supported on the [Message Read](read-scenario.md) surface. Support varies by client and Exchange environment, as shown in the following table.

| Client | Exchange Online | Exchange Subscription Edition (SE) | Exchange Server 2019 | Exchange Server 2016 |
| ------ | --------------- | ---------------------------------- | -------------------- | -------------------- |
| **Web browser** | Not available | Not available | Not available | Not available |
| **Windows (new)** | Not available | Not available | Not available | Not available |
| **Windows (classic)**<br>Version 2510 (Build 19312.20000) and later | In preview | Not available | Not available | Not available |
| **Mac** | Not available | Not available | Not available | Not available |
| **Android** | Not available | Not available | Not available | Not available |
| **iOS** | Not available | Not available | Not available | Not available |

#### Preview the decryption APIs in classic Outlook on Windows

To preview the decryption APIs in classic Outlook on Windows, join the [Microsoft 365 Insider program](https://aka.ms/MSFT365InsiderProgram), then choose the Beta Channel in the Outlook client. Your client must be on Version 2510 (Build 19312.20000) or later.

Classic Outlook on Windows includes a local copy of the production and beta versions of Office.js instead of loading from the content delivery network (CDN). By default, the local production copy of the API is referenced. To reference the local beta copy of the API, you must configure your computer's registry. This will enable you to test [preview features](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview) in your event handlers in classic Outlook on Windows.

1. In the registry, navigate to `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`. If the key doesn't exist, create it.

1. Create an entry named `EnableBetaAPIsInJavaScript` and set its value to `1`.

    :::image type="content" source="../images/outlook-beta-registry-key.png" alt-text="The `EnableBetaAPIsInJavaScript` key is set to 1 in the Registry Editor.":::

### Configure the manifest

> [!NOTE]
> The `OnMessageRead` event can currently only be implemented with an add-in only manifest.

To activate your add-in on the `OnMessageRead` event, you must configure the [\<VersionOverridesV1_1\>](/javascript/api/manifest/versionoverrides-1-1-mail) node of your add-in's **manifest.xml** file as follows.

- To run an event-based add-in in classic Outlook on Windows, you must specify the JavaScript file that contains the event handler in the [\<Override\>](/javascript/api/manifest/override) child element of the [\<Runtime\>](/javascript/api/manifest/runtime) element.
- Specify the `OnMessageRead` event in the `Type` attribute of a [\<LaunchEvent\>](/javascript/api/manifest/launchevent?view=outlook-js-preview&preserve-view=true) element. You must assign the function name of the event handler to the `FunctionName` attribute of the element. To facilitate checking whether the message was encrypted by the add-in, a header key must be specified in the `HeaderName` attribute. The same header is added to a message that's encrypted by the add-in.

The following is an example of a `<VersionOverrides>` node that implements the `OnMessageRead` event.

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
            <!-- References the HTML file that links to the event handler. -->
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
              <LaunchEvent Type="OnMessageRead" FunctionName="onMessageReadHandler" HeaderName="contoso-encrypted"/>
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

### Implement event handling

The `OnMessageRead` event handler is used to run the decryption operation and determine whether to display the decrypted contents of a message.

- To ensure your handler runs when the `OnMessageRead` event occurs, call `Office.actions.associate` in the JavaScript file where the handler is implemented. This maps the handler name specified in the `FunctionName` attribute of the `<LaunchEvent>` element in the manifest to its JavaScript counterpart.
- Once the decryption operation finishes, you must call `event.completed` to signal to the client that your add-in has completed processing the `OnMessageRead` event. To display the decrypted contents of a message and its attachments, pass a [MessageDecryptEventCompletedOptions](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true) object to the `event.completed` call and set its [allowEvent](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-messagedecrypteventcompletedoptions-allowevent-member) property to `true`. Then, specify the decrypted contents of the message in the object's [emailBody](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-messagedecrypteventcompletedoptions-emailbody-member) and [attachments](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-messagedecrypteventcompletedoptions-attachments-member) properties. You can also specify any data that your add-in may need for processing in the [contextData](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-messagedecrypteventcompletedoptions-contextdata-member) property. For example, you can store custom internet headers to decrypt messages in reply and forward scenarios. 

> [!NOTE]
> Be mindful of the following when creating an event-based add-in for classic Outlook on Windows.
>
> - Imports aren't currently supported in the JavaScript file containing the event handler.
> - When the JavaScript function specified in the manifest to handle an event runs, code in `Office.onReady()` and `Office.initialize` isn't run. We recommend adding any startup logic needed by the event handler, such as checking the user's Outlook version, to the event handler instead.

The following is an example of an `OnMessageRead` event handler.

```javascript
function onMessageReadHandler(event) {
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

    // Decrypted content and properties of a mail item.
    const decryptedEmailFile = "VGhpcyBpcyBhIHRleHQgZmlsZS4=...";
    const emailFileName = "Fabrikam_Report_202508.eml";

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
          attachmentType: Office.MailboxEnums.AttachmentType.Item,
            content: decryptedEmailFile,
            isInline: false,
            name: emailFileName
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
Office.actions.associate("onMessageReadHandler", onMessageReadHandler);
```

> [!TIP]
> In classic Outlook on Windows, when images are added to a message as inline attachments, they're automatically assigned a content ID. In the body of a message, the content ID of an inline attachment is specified in the `src` attribute of the `<img>` element similar to the following example.
>
> `<img width=96 height=96 id="Picture_1" src="cid:image001.png@01DC1E6F.FC7C7410">`
>
> To easily identify and provide these inline attachments during decryption, we recommend saving the content IDs of inline attachments to the message header during encryption. Call [Office.context.mailbox.item.getAttachmentsAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-getattachmentsasync-member(1)) to get the [content ID](/javascript/api/outlook/office.attachmentdetailscompose?view=outlook-js-preview&preserve-view=true#outlook-office-attachmentdetailscompose-contentid-member) of an inline attachment. Then, call [Office.context.mailbox.item.internetHeaders.setAsync](/javascript/api/outlook/office.internetheaders#outlook-office-internetheaders-setasync-member(1)) to save the ID to the header of the message.

## Behavior and limitations

- Be aware of the behaviors and limitations of event-based add-ins. To learn more, see [Activate add-ins with events](../develop/event-based-activation.md#behavior-and-limitations).
- Since each add-in uses its own encryption protocol, a message can only be decrypted by the same add-in that encrypted it. When a user doesn't have the required add-in installed to decrypt a message, a notification isn't shown to them. To guide the user through the decryption process, customize a placeholder message for the body of the encrypted message. The placeholder message can include information on how to install your add-in. To set the message body during the encryption process, call [Office.context.mailbox.item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(2)).

    :::image type="content" source="../images/outlook-encryption-placeholder-message.png" alt-text="A sample placeholder message of an encrypted message.":::

- To ensure data security and confidentiality, decrypted content isn't stored on the Outlook client. The contents of an encrypted message are decrypted every time a user opens it.
- An encrypted message must first be decrypted before a user can reply or forward it. A user can't reply or forward an encrypted message while it's being decrypted.
- If a user navigates to another mail item while an encrypted message is being decrypted, the decryption process stops running. The user must select or open the message again to activate the decryption process.
- When replying to or forwarding encrypted messages, drafts are saved unencrypted in the **Drafts** folder.

### Decryption notifications

Add-ins that handle the `OnMessageRead` event automatically display notifications in certain decryption scenarios as described in the following table.

| Notification | Scenario |
| ------------ | -------- |
| \<Add-in name\> is unavailable and can't process your message at this time | Applies to classic Outlook on Windows only. This notification is shown when the add-in fails to load because an error prevented the add-in from loading or the user's client or machine is offline. |
| \<Add-in name\> failed to process your message | An error was encountered while the add-in was decrypting the message. To retry the decryption operation, the recipient must switch to another message, then open the encrypted message again to invoke the `OnMessageRead` event. |
| \<Add-in name\> is processing your message | The add-in is handling the `OnMessageRead` event to decrypt the message. |
| \<Add-in name\> has finished processing your message | The add-in successfully decrypted the contents of the message. The user can now view the message and its attachments. |
| \<Add-in name\> is taking longer than expected to process your message | The add-in has been running for more than five seconds, but less than five minutes. |
| \<Add-in name\> timed out. To retry, select another email and then return to this message | The add-in times out after running for five minutes. To retry the decryption operation, the recipient must switch to another message, then open the encrypted message again to invoke the `OnMessageRead` event. |

## See also

- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
- [Activate add-ins with events](../develop/event-based-activation.md)
- [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md)
- [Get and set internet headers on a message in an Outlook add-in](internet-headers.md)
- [Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)
