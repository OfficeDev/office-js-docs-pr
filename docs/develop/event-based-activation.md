---
title: Activate add-ins with events
description: Learn how to develop an Office Add-in that implements event-based activation.
ms.date: 11/27/2025
ms.topic: concept-article
ms.localizationpriority: medium
---

# Activate add-ins with events

Event-based activation automatically triggers your add-in to complete their tasks without explicitly launching it. This allows the add-in to validate, insert, or refresh critical content without any manual operations. The add-in is opened in the background to avoid disrupting the user. You can also integrate event-based activation with the task pane and function commands.

## Overview

While the particular steps to add event-based functionality to your add-in vary by platform and manifest type, the general flow is as follows.

1. Update the manifest with an extension for the event.
1. Connect the event in the manifest with a JavaScript function to handle the event.
1. Have the event handler function perform its actions, then call `event.completed` when it finishes.
1. Call [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) to connect the event handler function with the ID specified in the manifest.

## Try out event-based activation

Discover how to streamline workflows and improve user experiences with event-based activation. Try out the samples to see the feature in action.

### Outlook samples

- [Automatically set the subject of a new message or appointment](../outlook/on-new-compose-events-walkthrough.md)
- [Automatically check for an attachment before a message is sent](../outlook/smart-alerts-onmessagesend-walkthrough.md)
- [Automatically update your signature when switching between mail accounts](../outlook/onmessagefromchanged-onappointmentfromchanged-events.md)
- [Encrypt attachments, process meeting request attendees, and react to appointment date/time changes using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
- [Set your signature using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
- [Identify and tag external recipients using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
- [Verify the color categories of a message or appointment before it's sent using Smart Alerts](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
- [Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label)

### Word samples

- [Add headers when a document opens](../word/add-headers-on-document-open.md)

## Supported events

The following tables list events that are currently available and the supported clients for each event. When an event is raised, the handler receives an `event` object which may include details specific to the type of event. The **Description** column includes a link to the related object where applicable.

### Excel, PowerPoint, Word events

| Event canonical name</br>and add-in only manifest name | Unified manifest for Microsoft 365 name | Description | Supported clients and channels |
| ----- | ----- | ----- | ----- |
| `OnDocumentOpened` | *Not yet supported* | Occurs when a user opens a document or creates a new document, spreadsheet, or presentation. | <ul><li>Windows (Build >= 16.0.18324.20032)</li><li>Office on the web</li><li>Office on Mac will be available later </li></ul>|

### Outlook events

Support for this feature in Outlook was introduced in [requirement set 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), with additional events now available in subsequent requirement sets. The following table lists each event's minimum requirement set and the clients and platforms that support it. For more information on Outlook clients and the requirement sets they support, see [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

|Event canonical name</br>and add-in only manifest name|Unified manifest for Microsoft 365 name|Description|Minimum requirement set and supported clients|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li><li>Android<sup>2</sup> <sup>3</sup></li><li>iOS<sup>2</sup> <sup>3</sup></li></ul>|
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|On creating a new appointment but not on editing an existing one.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|On adding or removing attachments while composing a message.<br><br>Event-specific data object: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|On adding or removing attachments while composing an appointment.<br><br>Event-specific data object: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|On adding or removing recipients while composing a message.<br><br>Event-specific data object: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li><li>Android<sup>2</sup> <sup>3</sup></li><li>iOS<sup>2</sup> <sup>3</sup></li></ul>|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|On adding or removing attendees while composing an appointment.<br><br>Event-specific data object: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|On changing date/time while composing an appointment.<br><br>Event-specific data object: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)<br><br>**Important**: If you drag and drop an appointment to a different date/time slot on the calendar, the `OnAppointmentTimeChanged` event doesn't occur. It only occurs when the date/time is directly changed from an appointment. |[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|On adding, changing, or removing the recurrence details while composing an appointment. If the date/time is changed, the `OnAppointmentTimeChanged` event also occurs.<br><br>Event-specific data object: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|On dismissing a notification while composing a message or appointment item. Only the add-in that added the notification will be notified.<br><br>Event-specific data object: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnMessageSend`|messageSending|On sending a message item. To learn more, try the [Smart Alerts walkthrough](../outlook/smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnAppointmentSend`|appointmentSending|On sending an appointment item. To learn more, see [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](../outlook/onmessagesend-onappointmentsend-events.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnMessageCompose`|messageComposeOpened|On composing a new message (includes reply, reply all, and forward) or editing a draft.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|On creating a new appointment or editing an existing one.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<ul><li>Web browser</li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnMessageFromChanged`|messageFromChanged|On changing the mail account in the **From** field of a message being composed. To learn more, see [Automatically update your signature when switching between Exchange accounts](../outlook/onmessagefromchanged-onappointmentfromchanged-events.md).|[1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13)<ul><li>Web browser<sup>4</sup></li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)<sup>4</sup> and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li><li>Android<sup>2</sup> <sup>3</sup></li><li>iOS<sup>2</sup> <sup>3</sup></li></ul>|
|`OnAppointmentFromChanged`|appointmentFromChanged|On changing the mail account in the organizer field of an appointment being composed. To learn more, see [Automatically update your signature when switching between Exchange accounts](../outlook/onmessagefromchanged-onappointmentfromchanged-events.md).|[1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13)<ul><li>New Mac UI<sup>2</sup></li></ul>|
|`OnSensitivityLabelChanged`|sensitivityLabelChanged|On changing the sensitivity label while composing a message or appointment. To learn how to manage the sensitivity label of a mail item, see [Manage the sensitivity label of your message or appointment in compose mode](../outlook/sensitivity-label.md).<br><br>Event-specific data object: [SensitivityLabelChangedEventArgs](/javascript/api/outlook/office.sensitivitylabelchangedeventargs?view=outlook-js-preview&preserve-view=true)|[1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13)<ul><li>Web browser<sup>4</sup></li><li>Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)<sup>4</sup> and classic<sup>1</sup>)</li><li>New Mac UI<sup>2</sup></li></ul>|
|`OnMessageReadWithCustomAttachment`|Not available|On opening a message that contains a specific attachment type in read mode.|[Preview](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<sup>5</sup><ul><li>Windows (classic<sup>1</sup>)</li></ul>|
|`OnMessageReadWithCustomHeader`|Not available|On opening a message that contains a specific internet header name in read mode.|[Preview](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<sup>5</sup><ul><li>Windows (classic<sup>1</sup>)</li></ul>|
|`OnMessageRead` (preview)|Not available|On matching the header of an encrypted message to the header key in an add-in's manifest. To learn more, see [Create an encryption Outlook add-in](../outlook/encryption-decryption.md).|[Preview](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<ul><li>Windows (classic<sup>1</sup>)</li></ul>|

> [!NOTE]
> <sup>1</sup> Event-based add-ins in classic Outlook on Windows require a minimum of Windows 10 Version 1903 (Build 18362) or Windows Server 2019 Version 1903 to run.
>
> <sup>2</sup> Add-ins that use the unified manifest for Microsoft 365 aren't directly supported in Outlook on Mac and on mobile devices. To run this type of add-in on Mac and on mobile platforms, the add-in must first be published to [Microsoft Marketplace](https://marketplace.microsoft.com/) then deployed in the [Microsoft 365 Admin Center](../publish/publish.md). For more information, see the "Client and platform support" section of [Office Add-ins with the unified app manifest for Microsoft 365](../develop/unified-manifest-overview.md#client-and-platform-support).
>
> <sup>3</sup> For more information, see [Implement event-based activation in Outlook mobile add-ins](../outlook/mobile-event-based.md).
>
> <sup>4</sup> The `OnMessageFromChanged` and `OnSensitivityLabelChanged` events aren't currently available with the unified manifest for Microsoft 365 in Outlook on the web and the new Outlook on Windows. To handle these events, implement an add-in only manifest instead. For information about the types of manifests, see [Office Add-ins manifest](add-in-manifests.md).
>
> <sup>5</sup> To preview the `OnMessageReadWithCustomAttachment` and `OnMessageReadWithCustomHeader` events, you must install classic Outlook on Windows Version 2312 (Build 17110.10000) or later. Then, join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/blog/microsoft365insiderblog/join-the-microsoft-365-insider-program-on-windows/4206638) and select the **Beta Channel** option to access Office beta builds.

#### Event-based activation in Outlook on mobile devices

Outlook on mobile supports APIs up to Mailbox requirement set 1.5. However, support is now enabled for additional APIs and features introduced in later requirement sets, such as the `OnNewMessageCompose` event. To learn more, see [Implement event-based activation in Outlook mobile add-ins](../outlook/mobile-event-based.md).

## Behavior and limitations

As you develop an event-based add-in, be mindful of the following feature behaviors and limitations.

- Event-based add-ins work only when deployed by an administrator. If users install them directly from Microsoft Marketplace or the Office Store, they will not automatically launch (for workarounds to the Microsoft Marketplace limitation, see [Microsoft Marketplace listing options for your event-based add-in](../publish/autolaunch-store-options.md)). Admin deployments are done by uploading the manifest to the Microsoft 365 admin center.

- APIs that interact with the UI or display UI elements are not supported for Word, PowerPoint, and Excel on Windows. This is because the event handler runs in a JavaScript-only runtime. For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).

- Event-based add-ins require an internet connection to be able to launch when a specific event occurs. Add-in event handlers are expected to be short-running, lightweight, and as noninvasive as possible. After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, your associated event handler must call the [event.completed](/javascript/api/outlook/office.mailboxevent#outlook-office-mailboxevent-completed-member(1)) method. (Note that code included after the `event.completed` statement isn't guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset. The add-in ends after it times out, or the user closes the compose window or sends the item.

- The behavior of multiple add-ins that subscribe to the same event isn't deterministic. Outlook launches the add-ins in no particular order. For Excel, PowerPoint, and Word, only one random add-in will be activated. For example, if multiple Word add-ins that handle `OnDocumentOpened`, only one of those handlers will run.

- Currently, only five event-based add-ins can be actively running.

- In all supported Outlook clients, the user must remain on the current mail item where the add-in was activated for it to complete running. Navigating away from the current item (for example, switching to another compose window or tab) terminates the add-in operation. However, an add-in that activates on the `OnMessageSend` event handles item switching differently depending on which Outlook client it's running on. To learn more, see the "User navigates away from current message" section of [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](../outlook/onmessagesend-onappointmentsend-events.md#user-navigates-away-from-current-message).

- In addition to item switching, an event-based add-in also ceases operation when the user sends the message or appointment they're composing.

### Event-based add-in limitations in Excel, PowerPoint, Word, and classic Outlook on Windows

When developing an event-based add-in to run on a Windows client, be mindful of the following:

- Imports aren't supported in the JavaScript file where you implement the handling for event-based activation.
- Only the JavaScript file referenced in the manifest is supported for event-based activation. You must bundle your event-handling JavaScript code into this single file. The location of the referenced JavaScript file in the manifest varies depending on the type of manifest your add-in uses.
  - **Add-in only manifest**: `<Override>` child element of the `<Runtime>` node
  - **Unified manifest for Microsoft 365**: `"script"` property of the `"code"` object

  Note that a large JavaScript bundle may cause issues with the performance of your add-in. We recommend preprocessing heavy operations, so that they're not included in your event-handling code.
- When the JavaScript function specified in the manifest to handle an event runs, code in `Office.onReady()` and `Office.initialize` isn't run. We recommend adding any startup logic needed by event handlers, such as checking the user's client version, to the event handlers instead.

### Event-based add-in limitations in Excel, PowerPoint, and Word

The following platforms or features are not yet supported.

- Office on Mac
- The unified manifest for Microsoft 365

### Event-based add-in limitations in Outlook on the web and the new Outlook on Windows

In Outlook on the web and the new Outlook on Windows, event-based activation is only supported on standard read and compose message and appointment surfaces. Event-based activation may not work when composing on some non-standard surfaces. For example:

- Responding to a meeting invite using the **RSVP with note** option.
- Forwarding a meeting from the calendar.

### Unsupported APIs

Some Office.js APIs that change or alter the UI aren't allowed from event-based add-ins. The following are blocked APIs.

| API | Methods |
| --- | --- |
| `Office.devicePermission` | <ul><li>`requestPermissionsAsync`</li></ul> |
| `Office.context.auth`\* | <ul><li>`getAccessToken`</li><li>`getAccessTokenAsync`</li></ul> |
| `Office.context.mailbox` | <ul><li>`displayAppointmentForm`</li><li>`displayMessageForm`</li><li>`displayNewAppointmentForm`</li><li>`displayNewMessageForm`</li></ul> |
| `Office.context.mailbox.item` | <ul><li>`close`</li></ul> |
| `Office.context.ui` | <ul><li>`displayDialogAsync`</li><li>`messageParent`</li></ul>|

> [!NOTE]
> \* [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) is supported in all versions that support event-based activation and single sign-on (SSO), while [Office.auth](/javascript/api/office/office.auth) is only supported in certain Outlook builds. For more information, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Outlook add-in](../develop/use-sso-in-event-based-activation.md).

### Preview features in event handlers (classic Outlook on Windows)

Classic Outlook on Windows includes a local copy of the production and beta versions of Office.js instead of loading from the content delivery network (CDN). By default, the local production copy of the API is referenced. To reference the local beta copy of the API, you must configure your computer's registry. This will enable you to test [preview features](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview) in your event handlers in classic Outlook on Windows.

1. In the registry, navigate to `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`. If the key doesn't exist, create it.
1. Create an entry named `EnableBetaAPIsInJavaScript` and set its value to `1`.

    :::image type="content" source="../images/outlook-beta-registry-key.png" alt-text="The EnableBetaAPIsInJavaScript registry value is set to 1.":::

## Enable single sign-on (SSO)

To enable SSO in your event-based add-in, you must add its JavaScript file to a well-known URI. For guidance on how to configure this resource, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Office Add-in](../develop/use-sso-in-event-based-activation.md).

### Request external data

You can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XMLHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

> [!NOTE]
> If your add-in will operate in a JavaScript-only runtime, use absolute URLs in your Fetch API calls. Relative URLs in Fetch API calls aren't supported in a JavaScript-only runtime.

Be aware that you must use additional security measures when using XMLHttpRequest objects, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and [CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS).

> [!NOTE]
> Full CORS support is available in Office on the web, Mac, and Windows (starting in Version 2201, Build 16.0.14813.10000) clients.

To make CORS requests from your event-based add-in, you must add the add-in and its JavaScript file to a well-known URI. For guidance on how to configure this resource, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Office Add-in](../develop/use-sso-in-event-based-activation.md).

## Troubleshoot your add-in

As you develop your event-based add-in, you may need to troubleshoot issues, such as your add-in not loading or the event not occurring. For guidance on how to troubleshoot an event-based add-in, see [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md).

## Deploy your add-in

Depending on the Office application, event-based add-ins can be deployed through one of the following options.

- **Admin-managed deployment**: Add-in is deployed through the Microsoft 365 admin center.
- **Restricted listing on Microsoft Marketplace**: Add-in is published to Microsoft Marketplace, but it doesn't appear in search results. Add-in acquisition requires a flight code URL. The add-in must still be deployed by an admin for the event-based activation feature to work.
- **Unrestricted listing on Microsoft Marketplace**: Add-in is published to Microsoft Marketplace and can be searched for by users and admins using the add-in's name or ID. Admin deployment isn't necessary for the event-based activation feature to work. The add-in must meet certain requirements for unrestricted listing.

The following table outlines the deployment options for event-based activation by Office application.

| Office application | Admin-managed deployment | Microsoft Marketplace |
| --- | --- | --- |
| **Excel** | Supported | Restricted listing option |
| **Outlook** | Supported | Restricted and unrestricted listing options |
| **PowerPoint** | Supported | Restricted listing option |
| **Word** | Supported | Restricted listing option |

For instructions on how to deploy an add-in through the Microsoft 365 admin center, see [Admin-managed deployment](#admin-managed-deployment). To learn more about listing your event-based add-in in Microsoft Marketplace, see [Microsoft Marketplace listing options for your event-based add-in](../publish/autolaunch-store-options.md).

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

### Admin-managed deployment

Admin deployments are done by uploading the manifest to the Microsoft 365 admin center. To do so, follow these steps.

1. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action.

:::image type="content" source="../images/outlook-deploy-event-based-add-ins.png" alt-text="The Integrated apps page on the Microsoft 365 admin center with the Upload custom apps action highlighted.":::

For more information about how to deploy an add-in, please refer to [Deploy and publish Office Add-ins in the Microsoft 365 admin center](/microsoft-365/admin/manage/office-addins).

### Deploy manifest updates

If an event-based add-in was admin-deployed, any change you make to the manifest requires admin consent through the Microsoft 365 admin center. Until the admin accepts your changes, users in their organization are blocked from using the add-in. To learn more about the admin consent process, see [Admin consent for installing event-based add-ins](../publish/autolaunch-store-options.md#admin-consent-for-installing-event-based-add-ins).

## See also

- [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md)
- [Debug event-based and spam-reporting add-ins](../testing/debug-autolaunch.md)
- [Microsoft Marketplace listing options for your event-based add-in](../publish/autolaunch-store-options.md)
- [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](../outlook/onmessagesend-onappointmentsend-events.md)
