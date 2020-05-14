---
title: Configure your add-in for event-based activation
description: Learn how to configure your add-in for event-based activation.
ms.topic: article
ms.date: 05/14/2020
localization_priority: Normal
---

# Configure your add-in for event-based activation (preview)

Before the introduction of the event-based activation feature, a user would have to explicitly launch an add-in to complete their tasks. This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item. You can also integrate with the task pane and UI-less functionality. At present, the supported events are as follows.

- Compose a new message
- Organize a new appointment

> [!IMPORTANT]
> This feature is only supported in preview in Outlook on the web with an Office 365 subscription. See the [How to preview](#how-to-preview) section later in this article for more details.
>
> Because preview APIs are subject to change without notice, they shouldn't be used in production add-ins.

## Configure the manifest

To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest. For now, `DesktopFormFactor` is the only supported form factor.

> [!TIP]
> To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).

The following example shows an excerpt from the manifest that runs the add-in's event-based activation functionality whenever a new message is created.

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <Runtimes>
          <Runtime resid="runtime1">
            <Override type="javascript" resid="runtimeJs"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- Configure AppointmentOrganizerCommandSurface extension point to support
          add-in activation on composing a new appointment. -->

          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="OnMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="OnAppointmentComposeHandler"/>
            </LaunchEvents>
            <SourceLocation resid="runtime1"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file. You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook endpoint. As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.

## Event-based activation behavior and limitations

Add-ins that activate based on events are designed to be short-running, up to 330 seconds only. We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event. The add-in is also ended when the user closes the compose window.

If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order. Currently, only 5 event-based add-ins can be actively running. Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.

The user can switch or navigate away from the item (that is, moves away from current mail item) where the add-in started running. The add-in that was launched will finish its operation in the background.

Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.

- Under `Office.context.mailbox`:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Under `Office.context.ui`:
  - `displayDialogAsync`
  - `messageParent`
- Under `Office.context.auth`:
  - `getAccessTokenAsync`

## How to preview

We invite you to try out the event-based activation feature! Then let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).

To preview this feature:

- Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview`.
- You may need to join the [Office Insider program](https://insider.office.com) for access to more recent Office builds.
- Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).

## See also

[Outlook add-in manifests](manifests.md)