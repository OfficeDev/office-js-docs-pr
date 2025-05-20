---
title: Special requirements for add-ins on the iPad
description: Learn some requirements for creating an Office Add-in that runs on an iPad.
ms.topic: best-practice
ms.date: 08/15/2024
ms.localizationpriority: medium
---

# Special requirements for add-ins on the iPad

If your add-in uses only Office APIs that are supported on the iPad, then customers can install it on iPads (see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md) for more information.) If the add-in is published in [AppSource](https://appsource.microsoft.com), then there are some requirements for add-ins that can be installed on iPads, in addition to [the best practices that apply to all Office Add-ins](../concepts/add-in-development-best-practices.md).

The following table lists the tasks to perform.

|Task|Description|Resources|
|:-----|:-----|:-----|
|Apply iOS design best practices.|Integrate your add-in UI seamlessly with the iOS experience.| See note following this table. |
|Optimize your add-in for touch.|Make your UI responsive to touch inputs in addition to mouse and keyboard.|[Apply UX design principles](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Make your add-in free.|Office on iPad is a channel through which you can reach more users and promote your services. These new users have the potential to become your customers.|[Certification policy 1120.2](/legal/marketplace/certification-policies#11202-mobile-requirements)|
|Make your add-in commerce free on the iPad.|When it's running on the iPad, your add-in must be free of in-app purchases, trial offers, UI that aims to upsell to a non-free version, or links to any online stores where users can purchase or acquire other content, apps, or add-ins. Your Privacy Policy and Terms of Use pages must also be free of any commerce UI or AppSource links.|[Certification policy 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Your add-in can still have commerce on other platforms. To do so, test the [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member) property and suppress all commerce when it returns `false`.|
|Submit your add-in to AppSource.|In Partner Center, on the **Product setup** page, select the **Make my product available on iOS and Android (if applicable)** check box, and provide your Apple developer ID in Account settings. Review the [Application Provider Agreement](https://go.microsoft.com/fwlink/?linkid=715691) to make sure you understand the terms.|[Make your solutions available in AppSource and within Office](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Your add-in can serve an alternate UI based on the device that it's running on. To detect whether your add-in is running on an iPad, you can use the following APIs.
>
> - const isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member)
> - const allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member)
>
> On an iPad, `touchEnabled` returns `true` and `commerceAllowed` returns `false`.
>
> For information on the best UI design practices for iPad, see [Designing for iOS](https://developer.apple.com/design/human-interface-guidelines/designing-for-ios).

## Best practices for developing Office Add-ins that can run on iPad

Apply the following best practices for developing add-ins that run on iPad.

- **Develop and debug the add-in on Windows or Mac and sideload it to an iPad.**

    You can't develop the add-in directly on an iPad, but you can develop and debug it on a Windows or Mac computer and sideload it to an iPad for testing. Because an add-in that runs in Office on iOS or Mac supports the same APIs as an add-in running in Office on Windows, your add-in's code should run the same way on these platforms. For details, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md) and [Sideload Office Add-ins on iPad for testing](../testing/sideload-an-office-add-in-on-ipad.md).

- **Specify API requirements in your add-in's manifest or with runtime checks.**

    When you specify API requirements in your add-in's manifest, Office will determine if the Office client application supports those API members. If the API members are available in the application, then your add-in will be available. Alternatively, you can perform a runtime check to determine if a method is available in the application before using it in your add-in. Runtime checks ensure that your add-in is always available in the application, and provides additional functionality if the methods are available. For more information, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).

## Outlook add-ins on iPad

For information about designing Outlook add-ins that look good and work well in Outlook on mobile devices, see [Add-ins for Outlook on mobile devices](../outlook/outlook-mobile-addins.md).
