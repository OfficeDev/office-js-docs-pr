---
title: Configure your add-in to autolaunch
description: Learn how to configure your add-in to autolaunch.
ms.topic: article
ms.date: 05/12/2020
localization_priority: Normal
---

# Configure your add-in to autolaunch (preview)

Before the introduction of the autolaunch feature, a user would have to explicitly launch an add-in to complete their tasks. The autolaunch feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item. You can also integrate with the task pane and UI-less functionality. At present, the supported events are as follows.

- Compose a new message
- Organize a new appointment

> [!NOTE]
> This feature is only supported in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with an Office 365 subscription.

## Configure the manifest

To enable autolaunch scenarios in your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest. For now, `DesktopFormFactor` is the only supported form factor.

> [!TIP]
> To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).

The following example shows an excerpt from the manifest that runs the add-in's autolaunch functionality whenever a new message is created.

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
          autolaunch on composing a new appointment. -->

          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="SubjectChange"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="SubjectChange"/>
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

The `Runtime` element provides the location of the HTML and JavaScript for autolaunch. Outlook on Windows uses the JavaScript file, while Outlook on the web uses the HTML that references the same JavaScript file. However, you must provide both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook endpoint.  

## Autolaunch behavior and limitations

Autolaunch add-ins are designed to be short-running, for up to 330 seconds. We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event. The add-in is also ended when the user closes the compose window.

If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order. Currently, only 5 autolaunch add-ins can be actively running. Any additional autolaunch add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.

The user can switch or navigate away from the item (that is, moves away from current mail item) where the autolaunch add-in started running. The add-in that was launched will finish its operation in the background.

Some Office.js APIs that change or alter the UI are not allowed from autolaunch add-ins. The following are the blocked APIs.

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

## See also

[Outlook add-in manifests](manifests.md)