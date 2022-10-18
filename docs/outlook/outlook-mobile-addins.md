---
title: Outlook add-ins for Outlook Mobile
description: Outlook mobile add-ins are supported on all Microsoft 365 business accounts and Outlook.com accounts.
ms.date: 10/17/2022
ms.localizationpriority: medium
---

# Add-ins for Outlook Mobile

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Outlook mobile add-ins are supported on all Microsoft 365 business accounts and Outlook.com accounts. However, support is not currently available on Gmail accounts.

**An example task pane in Outlook on iOS**

![Screenshot of a task pane in Outlook on iOS.](../images/outlook-mobile-addin-taskpane.png)

<br/>

**An example task pane in Outlook on Android**

![Screenshot of a task pane in Outlook on Android.](../images/outlook-mobile-addin-taskpane-android.png)

## What's different on mobile?

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
  - The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).
  - The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

- In general, only Message Read mode is supported at this time. That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest. However, there are a couple of exceptions:
  1. Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface). See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.
  1. Appointment Attendee mode is supported for integrated add-ins created by providers of note-taking and customer relationship management (CRM) applications. Such add-ins should instead declare the [MobileLogEventAppointmentAttendee extension point](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee). See the [Log appointment notes to an external application in Outlook mobile add-ins](mobile-log-appointments.md) article for more about this scenario.

- The [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- When you submit your add-in to the store with [MobileFormFactor](/javascript/api/manifest/mobileformfactor) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.

- Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](/javascript/api/manifest/control) and [icon sizes](/javascript/api/manifest/icon) included.

## What makes a good scenario for mobile add-ins?

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Here are examples of scenarios that make sense in Outlook Mobile.

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**An example user interaction to create a Trello card from an email message on iOS**

![Animated GIF showing user interaction with an Outlook Mobile add-in on iOS.](../images/outlook-mobile-addin-interaction.gif)

<br/>

**An example user interaction to create a Trello card from an email message on Android**

![Animated GIF showing user interaction with an Outlook Mobile add-in on Android.](../images/outlook-mobile-addin-interaction-android.gif)

## Testing your add-ins on mobile

To test an add-in on Outlook Mobile, first [sideload an add-in](sideload-outlook-add-ins-for-testing.md) to a Microsoft 365 or Outlook.com account on the web, Windows, or Mac. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load in your Outlook client on mobile.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

Troubleshooting on mobile can be hard since you may not have the tools you're used to. However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

> [!NOTE]
> Modern Outlook on the web on iPhone and Android smartphones is no longer required or available for testing Outlook add-ins. Additionally, add-ins aren't supported in Outlook on Android, on iOS, and modern mobile web with on-premises Exchange accounts. Certain iOS devices still support add-ins when using on-premises Exchange accounts with classic Outlook on the web. For information about supported devices, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet).

## Next steps

Learn how to:

- [Add mobile support to your add-in's manifest](add-mobile-support.md).
- [Design a great mobile experience for your add-in](outlook-addin-design.md).
- [Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.
