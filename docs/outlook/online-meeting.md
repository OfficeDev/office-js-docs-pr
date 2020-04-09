---
title: How to create an Outlook mobile add-in for an online-meeting provider (preview)
description: Discusses how to set up an Outlook mobile add-in for an online-meeting service provider.
ms.topic: article
ms.date: 04/09/2020
localization_priority: Normal
---

# Create an Outlook mobile add-in for an online-meeting provider (preview)

Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile. However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome. By implementing this feature, service providers can make the online meeting creation experience seamless for their Outlook add-in users.

> [!NOTE]
> This feature is only supported in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.

## Scenario: Online-meeting provider

In this section, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service. Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".

### Configure the manifest

To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`. Other form factors are not supported.

The following example shows a sample of the manifest that includes the `MobileFormFactor` element and `MobileOnlineMeetingCommandSurface` extension point.

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

### Implementation

In this section, you'll see an outline for how your add-in script can update a user's meeting to include online meeting details.

The following example shows how you could construct online meeting details. Not shown is how to get the meeting organizer's ID and other details from your service.

```js
const newBody = '<br>' +
    '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
    '<br><br>' +
    'Phone Dial-in: +1(123)456-7890' +
    '<br><br>' +
    'Meeting ID: 123 456 789' +
    '<br><br>' +
    'Want to test your video connection?' +
    '<br><br>' +
    '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
    '<br><br>';
```

The next example shows how you could define a UI-less function named `insertContosoMeeting` referenced in the manifest to update the meeting body with the online meeting details.

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

function insertContosoMeeting(event) {
    // Get HTML body from the client.
    mailboxItem.body.getAsync("html",
        { asyncContext: event },
        function (getBodyResult) {
            if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                updateBody(getBodyResult.asyncContext, getBodyResult.value);
            } else {
                console.error("Failed to get HTML body.");
                getBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
```

This sample shows an implementation of the supporting function `updateBody` used in the previous example that appends the online meeting details to the current body of the meeting.

```js
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
```

### Testing and validation

You can follow the usual guidance to [test and validate your add-in](testing-and-tips.md). After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, you should restart Outlook on your mobile device (Android is the only supported client for now) then verify on a new meeting screen that the Microsoft Teams or Skype toggle has been replaced with your own.

#### Create meeting UI

As a meeting organizer, you should see screens similar to the following when you create a meeting.

> [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png)

> [![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png)

> [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png)

> ![operands](../images/operands.png)

#### Join meeting UI

As a meeting attendee, you should see a screen similar to the following when you view the meeting.

> [![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view.png)](../images/outlook-android-join-online-meeting-view-expanded.png)

## Available APIs

***TODO:*** Confirm set of APIs - will need to update ref docs for each of them - point to this article
e.g. **Note**: This member is not supported in Outlook on iOS or Android. However, there's an exception for online-meeting providers. See *online meeting article* for details.

***TODO:*** Note similar exception in [mobile add-ins support](add-mobile-support.md#compose-mode-and-appointments) article.

The following APIs are available for this feature.

- [Dialog APIs for auth flow](../develop/dialog-api-in-office-add-ins.md)
- Appointment Organizer APIs
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))
  - [Office.context.mailbox.item.organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) ([Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-preview))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))
  - [Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))

## Restrictions

Several restrictions apply.

- Applicable only to online-meeting service providers.
- Currently in preview.
- At present, Android is the only supported client. Support on iOS is coming soon.
- Only admin-installed add-ins should show up on the meeting compose screen, replacing the default Teams or Skype option. User-installed add-ins won't activate.
- The add-in icon should be in grayscale using hex code #919191 or its equivalent in other color formats.
- Only one UI-less command is supported in Appointment Organizer (compose) mode.

## See also

- [Add-ins for Outlook Mobile](outlook-mobile-addins.md)
- [Add support for add-in commands for Outlook Mobile](add-mobile-support.md)