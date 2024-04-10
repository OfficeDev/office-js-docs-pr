---
title: Add-ins for Outlook on mobile devices
description: Outlook mobile add-ins are supported on all Microsoft 365 business accounts and Outlook.com accounts.
ms.date: 04/12/2024
ms.localizationpriority: medium
---

# Add-ins for Outlook on mobile devices

Add-ins now work in Outlook on mobile devices, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook mobile.

Outlook mobile add-ins are supported on all Microsoft 365 business accounts and Outlook.com accounts. However, support is not currently available on Gmail accounts.

**An example task pane in Outlook on iOS**

![A sample task pane in Outlook on iOS.](../images/outlook-mobile-addin-taskpane.png)

<br/>

**An example task pane in Outlook on Android**

![A sample task pane in Outlook on Android.](../images/outlook-mobile-addin-taskpane-android.png)

## What's different on mobile?

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for customers, any add-in declaring mobile support must meet certain validation criteria to be approved in AppSource.
  - The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).
  - The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-outlook-mobile-add-ins).
  - You'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.

- In general, only Message Read mode is supported. This has implications for how you configure the manifest.
  - **Unified manifest for Microsoft 365**: "mailRead" is the only item you should declare in the "extensions.ribbons.contexts" array.
  - **XML manifest**: `MobileMessageReadCommandSurface` is the only [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest. 
  
  However, there are some exceptions.
  1. Appointment Organizer mode is supported for online meeting provider integrated add-ins.
     - **Unified manifest for Microsoft 365**: "onlineMeetingDetailsOrganizer" is permitted in the "extensions.ribbons.contexts" array.
     - **XML manifest**: The [MobileOnlineMeetingCommandSurface extension point](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) is permitted. 
  
     For more information on this scenario, see [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md).

  1. Appointment Attendee mode is supported for integrated add-ins created by providers of note-taking and customer relationship management (CRM) applications. 
     - **Unified manifest for Microsoft 365**: "logEventMeetingDetailsAttendee" is permitted in the "extensions.ribbons.contexts" array.
     - **XML manifest**: The [MobileLogEventAppointmentAttendee extension point](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) is permitted. 
    
     For more information on this scenario, see [Log appointment notes to an external application in Outlook mobile add-ins](mobile-log-appointments.md).

  1. Event-based add-ins that activate on the `OnNewMessageCompose` event require an exception.
     - **Unified manifest for Microsoft 365**: Event-based add-ins aren't treated as a context in the unified manifest, so there is no exception for configuring the "extensions.ribbons.contexts" array. But note that event-based add-ins do require an "extensions.autoRunEvents" property in the manifest. 
     - **XML manifest**: The [LaunchEvent extension point](/javascript/api/manifest/extensionpoint#launchevent) *must be declared*. 
  
     For more information, see [Implement event-based activation in Outlook mobile add-ins](mobile-event-based.md).

- The [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API isn't supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- Your manifest needs to declare mobile support including special mobile controls and icon sizes. 
  - **Unified manifest for Microsoft 365**: Include the string "mobile" in the "extensions.ribbons.requirements.formFactors" array, and include a "customMobileGroup" property in the tab object in the "extensions.ribbons.tabs" array. This property must include a "buttonType" of "MobileButton" and an "icons" array.
  - **XML manifest**: Include a **\<MobileFormFactor\>**, and include the correct types of [controls](/javascript/api/manifest/control) and [icon sizes](/javascript/api/manifest/icon).
  
  To learn more, see [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).

## What makes a good scenario for Outlook mobile add-ins?

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Here are examples of scenarios that make sense in Outlook mobile.

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. For example, a customer relationship management (CRM) add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. For example, an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**An example user interaction to create a Trello card from an email message on iOS**

![Animated GIF showing user interaction with an add-in in Outlook on iOS.](../images/outlook-mobile-addin-interaction.gif)

<br/>

**An example user interaction to create a Trello card from an email message on Android**

![Animated GIF showing user interaction with an add-in in Outlook on Android.](../images/outlook-mobile-addin-interaction-android.gif)

## Testing your add-ins on mobile

To test an add-in on Outlook mobile, first [sideload an add-in](sideload-outlook-add-ins-for-testing.md) using a Microsoft 365 or Outlook.com account in Outlook on the web, on Windows, or on Mac. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load in Outlook mobile.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

Troubleshooting on mobile can be hard since you may not have the tools you're used to. However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).

> [!NOTE]
> Modern Outlook on the web on iPhone and Android smartphones is no longer required or available for testing Outlook add-ins. Additionally, add-ins aren't supported in Outlook on Android, on iOS, and modern mobile web with on-premises Exchange accounts. Certain iOS devices still support add-ins when using on-premises Exchange accounts with classic Outlook on the web. For information about supported devices, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet).

## Next steps

Learn how to:

- [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).
- [Implement supported Outlook JavaScript APIs](outlook-mobile-apis.md).
- [Design a great mobile experience for your add-in](outlook-addin-design.md).
- [Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.
