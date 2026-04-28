---
title: Test your Outlook add-in on mobile devices
description: Learn how to run and test your Outlook add-in on mobile device platforms, such as Android and iOS.
ms.date: 04/28/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Test your Outlook add-in on mobile devices

Testing an Outlook add-in on a mobile device requires a slightly different workflow than testing on the web or desktop. You don't sideload the add-in directly in Outlook on Android or on iOS. Instead, you first sideload it in Outlook on the web, on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) or classic), or on Mac, then validate the mobile experience on the devices you support.

Use the guidance in this article to set up a testing environment for mobile support and test an add-in in Outlook on mobile.

## Before you begin

To test your add-in in Outlook on mobile, you must meet the following prerequisites.

- Your add-in uses an [add-in only manifest](../develop/add-in-manifests.md).

    > [!NOTE]
    > Add-ins that use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md) aren't supported in Outlook on mobile devices. We're working hard to provide support in mobile devices. For more information, see [Support for add-ins with the unified manifest for Microsoft 365](outlook-mobile-addins.md#support-for-add-ins-with-the-unified-manifest-for-microsoft-365).

- Your manifest is configured for mobile support. For more information, see [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).
- You have a Microsoft 365 or Outlook.com account that supports add-ins in Outlook on mobile.

    [!INCLUDE [outlook-mobile-on-premises](../includes/outlook-mobile-on-premises.md)]
- Your add-in is hosted on an HTTPS endpoint that is reachable from the mobile device.

## Sideload the add-in

To make an add-in available in Outlook on mobile, sideload it to the same mailbox from Outlook on the web, on Windows, or on Mac.

1. Sideload the add-in by following [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).
1. Sign in to the Outlook app on your mobile device with the same Microsoft 365 or Outlook.com account that you used to sideload the add-in.
1. Open a supported item that matches the activation scenario in your manifest, such as a message in read mode.
1. Locate and open the add-in to test it. To locate an add-in from a message in Read mode, select **More options** (three vertical dots), then choose the add-in you want to test.

    :::image type="content" source="../images/outlook-mobile-add-ins.png" alt-text="The 'More options' menu that displays the installed add-ins.":::

## Remove a sideloaded add-in

To remove a sideloaded add-in from Outlook on mobile, remove it from Outlook on the web, on Windows (new or classic), or on Mac. For guidance, see [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).

> [!TIP]
> In Outlook on mobile, you can manage available and installed add-ins from **Settings** > **Add-ins**. A checkmark next to an add-in indicates that it's installed. If you clear the checkmark, the add-in no longer appears in the applicable activation surface (for example, in the **More options** menu of a message in Read mode). However, the add-in remains in the list of add-ins you can install. When you remove a sideloaded add-in during testing, the add-in is removed from the list.

## Troubleshoot during testing

If you need to troubleshoot your add-in during testing, try the following steps.

- Confirm that the add-in is installed for the correct mailbox account. For guidance on managing accounts in Outlook on Android or on iOS, see [Outlook for iOS & Android FAQs](https://support.microsoft.com/office/820ab229-529c-4e41-b227-b5dda358bfd4).
- Verify that the installed add-in appears for the same mailbox account in other supported clients, such as Outlook on the web.

    > [!NOTE]
    > Modern Outlook on the web on iPhone and Android smartphones is no longer available for testing Outlook add-ins.
- Verify that the manifest includes the required mobile configuration in the manifest. For information on supporting add-in commands on mobile, see [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).
- Confirm that the scenario is supported on mobile. For more information on supported modes and APIs, see [What makes a good scenario for Outlook mobile add-ins?](outlook-mobile-addins.md#what-makes-a-good-scenario-for-outlook-mobile-add-ins) and [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md).
- Ensure that the endpoint hosting the add-in is running and reachable from the mobile device.
- Use your preferred web debugging tool to send logs from your mobile device to an accessible endpoint for analysis.

    > [!NOTE]
    > Microsoft Edge DevTools isn't supported in Outlook on mobile devices.

## See also

- [Add-ins for Outlook on mobile devices](outlook-mobile-addins.md)
- [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md)
- [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md)
- [Design add-ins for Outlook on mobile devices](outlook-addin-design.md)
- [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md)
