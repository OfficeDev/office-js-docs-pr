---
title: Log appointments to a CRM application in Outlook mobile add-ins
description: Learn how to set up an Outlook mobile add-in to log appointments to a CRM application.
ms.topic: article
ms.date: 11/02/2021
ms.localizationpriority: medium
---

# Log appointments to a CRM application in Outlook mobile add-ins

Saving your appointment details and related notes to a CRM or note-taking application can help you keep track of meetings you've attended.

In this article, you'll learn how to set up your Outlook mobile add-in to enable users to log details and notes about their appointments to your CRM or note-taking application. Throughout this article, we'll be using a fictional service provider, "Contoso".

> [!IMPORTANT]
> This feature is only supported on Android with a Microsoft 365 subscription.

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.

## Configure the manifest

To enable users to log appointment notes with your add-in, you must configure the [MobileLogEventAppointmentAttendee extension point](../reference/manifest/extensionpoint.md#mobilelogeventappointmentattendee) in the manifest under the parent element `MobileFormFactor`. Other form factors are not supported.

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
            <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>logAppointmentDetails</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Log appointment details and notes to Contoso CRM"/>
        <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Log details and notes to Contoso CRM for this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).

## Implement capturing appointment details

In this section, learn how your add-in can extract appointment details when the user taps the **Log** button.

1. From the same quick start project, open the file **./src/commands/commands.js** in your code editor.

1. Replace the entire content of the **commands.js** file with the following JavaScript.

    ```js
    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    /* TODO: To update!!! */
    /* Code code code */

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.logAppointmentDetails = logAppointmentDetails;
    ```

## Implement viewing appointment details in a task pane

In this section, learn how to display the logged appointment details in a task pane when the user taps the **View** button.

```js
/* TODO: To update!!! */
/* Code code code */
```

## Implement viewing appointment details in a dialog

Alternatively, you can display the logged appointment details in a dialog when the user taps the **View** button. This section describes that implementation.

```js
/* TODO: To update!!! */
/* Code code code */
```

## Implement deleting a log

To delete the appointment log, use Microsoft Graph to clear the customer properties object when the user taps **Delete log**.

```js
/* TODO: To update!!! */
/* Code code code */
```

## Testing and validation

- Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).
- After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device.
- Open an appointment as an attendee then verify that under the Meeting Insights card, there's a new card with your add-in's name alongside the **Log** button.

### UI: Log appointment details

As a meeting attendee, you should see a screen similar to the following image when you open a meeting.

![Screenshot showing the Log button on an appointment screen on Android.](../images/outlook-android-log-appointment-details.png)

### UI: View appointment log

After successfully logging the appointment details, the button should now display the name **View** instead of **Log**. You should see a screen similar to the following image.

![Screenshot showing the View button on an appointment screen on Android.](../images/outlook-android-view-appointment-log.png)

## Available APIs

The following APIs are available for this feature.

- [Dialog APIs](../develop/dialog-api-in-office-add-ins.md)
- [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event?view=outlook-js-preview&preserve-view=true)
- [Office.RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- [Appointment Read (attendee) APIs](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true) **except** the following:
  - [Office.context.mailbox.item.categories](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories)
  - [Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation)
  - [Office.context.mailbox.item.isAllDayEvent](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent)
  - [Office.context.mailbox.item.recurrence](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence)
  - [Office.context.mailbox.item.sensitivity](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity)
  - [Office.context.mailbox.item.seriesId](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId)

## Restrictions

Several restrictions apply.

- The **Log** action button name cannot be changed.
- The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).
- The add-in should extract the meeting details from the appointment form within the one-minute timeout period. However, any time spent in a dialog box the add-in opened for authentication, etc. is excluded from the timeout period.

## See also

- [Add-ins for Outlook Mobile](outlook-mobile-addins.md)
- [Add support for add-in commands for Outlook Mobile](add-mobile-support.md)
