---
title: Best practices for developing Office Add-ins
description: Apply the best practices when developing Office Add-ins.
ms.topic: best-practice
ms.date: 07/28/2025
ms.localizationpriority: medium
---

# Best practices for developing Office Add-ins

Great add-ins provide unique, compelling functionality that extend Office apps in visually appealing ways. To build a successful add-in, you'll need to create an engaging first-time user experience, design a polished UI, and optimize performance. Follow the best practices in this article to help your users complete tasks quickly and efficiently.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## Provide clear value

Build add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office apps, such as:

- Make core authoring tasks faster and easier with fewer interruptions.
- Enable new scenarios within Office.
- Embed complementary services within Office apps.
- Improve the Office experience to enhance productivity.

Make sure users understand your add-in's value immediately by [creating an engaging first-run experience](#create-an-engaging-first-run-experience).

When you're ready to promote your add-in, learn how to create an [effective AppSource listing](/partner-center/marketplace-offers/create-effective-office-store-listings).

- Make your add-in's benefits clear in the title and description. Don't rely only on your brand to communicate what your add-in does.
- Ensure your add-in provides sufficient value to justify users' investment. It shouldn't be just a simple utility or have limited scope.
- [!INCLUDE [AppSource enterprise info](../includes/appsource-enterprise.md)]

## Create an engaging first-run experience

New users are still deciding whether to use or abandon your add-in after downloading it from the store. Here's how to win them over.

- **Make the next steps clear.** Use videos, placemats, paging panels, or other resources to guide users through your add-in.

- **Lead with value, not registration.** Reinforce your add-in's value proposition when it launches rather than immediately asking users to sign in.

- **Provide helpful guidance.** Include teaching UI to guide users and make the experience feel personal.

  ![A "Do" versus "Don't" comparison on how to guide your users to use the UI. The "Do" example shows an add-in that includes a button users can click to get started. The "Don't" example shows an add-in with no introductory steps or buttons.](../images/contoso-part-catalog-do-dont.png)

- **Show users what to expect.** If your content add-in binds to data in the user's document, include sample data or a template to show users the expected data format.

  ![A "Do" versus "Don't" comparison on including an option to insert sample data in your add-in. The "Do" example shows an add-in that includes a button users can click to insert sample data. The "Don't" example shows an add-in without sample data or buttons.](../images/add-in-title.png)

- **Offer free trials.** If your add-in requires a subscription, make some functionality available without one.

- **Simplify sign-up.** Prefill information like email and display name, and skip email verifications when possible.

- **Avoid pop-ups.** If you must use them, guide users on how to enable your pop-up window.

For patterns you can apply when developing your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).

## Use add-in commands

Provide relevant UI entry points for your add-in by using add-in commands. These commands help users discover and access your add-in's functionality directly from the Office ribbon. For details and design best practices, see [add-in commands](../design/add-in-commands.md).

## Apply UX design principles

Follow these key principles to create add-ins that feel native to Office:

- **Match the Office experience.** Ensure your add-in's look, feel, and functionality complement the Office experience. See [Design the UI of Office Add-ins](../design/add-in-design.md).

- **Prioritize content over chrome.** Avoid unnecessary UI elements that don't add value to the user experience.

- **Keep users in control.** Make sure users understand important decisions and can easily reverse actions your add-in performs.

- **Use branding thoughtfully.** Inspire trust and help orient users, but don't overwhelm or advertise to them.

- **Minimize scrolling.** Optimize for 1366 x 768 resolution.

- **Use licensed images only.** Avoid legal and branding issues that come from unlicensed images.

- **Write clearly.** Use [clear and simple language](../design/voice-guidelines.md) in your add-in.

- **Design for accessibility.** Make your add-in easy for all users to interact with and accommodate assistive technologies like screen readers. See our [accessibility guidelines](../design/accessibility-guidelines.md).

- **Support all platforms and input methods.** Design for mouse/keyboard and [touch](#optimize-for-touch). Ensure your UI responds well to different form factors.

### Optimize for touch

Touch support is essential for modern Office add-ins.

- **Detect touch support.** Use the [Context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member) property to detect whether the Office app your add-in runs on is touch enabled.

  > [!NOTE]
  > This property isn't supported in Outlook.

- **Size controls appropriately.** Make sure all controls work well with touch interaction. For example, buttons need adequate touch targets, and input boxes should be large enough for users to enter text easily.

- **Don't rely on hover or right-click.** These input methods aren't available on touch devices.

- **Support both orientations.** Ensure your add-in works in both portrait and landscape modes. Remember that on touch devices, the soft keyboard might hide part of your add-in.

- **Test on real devices.** Use [sideloading](../testing/sideload-an-office-add-in-on-ipad.md) to test your add-in on actual touch devices.

## Optimize and monitor add-in performance

Performance directly impacts user satisfaction. Follow these guidelines to keep your add-in fast and responsive:

- **Aim for quick loading.** Your add-in should load in 500 ms or less to create the perception of fast UI responses.

- **Respond quickly to interactions.** All user interactions should respond in under one second.

- **Show progress for long operations.** Provide loading indicators for operations that take time.

- **Use a CDN.** Host images, resources, and common libraries on a content delivery network (CDN). Load as much as possible from one place.

- **Follow web optimization best practices.** In production, use only minified versions of libraries. Load only the resources you need and optimize how they're loaded.

- **Provide feedback for longer operations.** When operations take time to execute, give users feedback based on the thresholds in the following table. For more information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).

  |Interaction class|Target|Upper bound|Human perception|
  |:-----|:-----|:-----|:-----|
  |Instant|<=50 ms|100 ms|No noticeable delay.|
  |Fast|50-100 ms|200 ms|Minimally noticeable delay. No feedback necessary.|
  |Typical|100-300 ms|500 ms|Quick, but too slow to be described as fast. No feedback necessary.|
  |Responsive|300-500 ms|1 second|Not fast, but still feels responsive. No feedback necessary.|
  |Continuous|>500 ms|5 seconds|Medium wait, no longer feels responsive. Might need feedback.|
  |Captive|>500 ms|10 seconds|Long, but not long enough to do something else. Might need feedback.|
  |Extended|>500 ms|>10 seconds|Long enough to do something else while waiting. Might need feedback.|
  |Long running|>5 seconds|>1 minute|Users will certainly do something else.|

- **Monitor your service.** Use telemetry to monitor service health and user success.

- **Minimize data exchanges.** Reduce data exchanges between your add-in and the Office document. For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).

## Special requirements for iPad add-ins

If your add-in uses only Office APIs that are supported on iPad, customers can install it on iPads. (See [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md) for more information.) If you're publishing to [AppSource](https://appsource.microsoft.com), there are additional requirements for add-ins that can be installed on iPads.

### iPad AppSource requirements

|Task|Description|Resources|
|:-----|:-----|:-----|
|Apply iOS design best practices.|Integrate your add-in UI seamlessly with the iOS experience.|[Designing for iOS](https://developer.apple.com/design/human-interface-guidelines/designing-for-ios)|
|Make your add-in free.|Office on iPad is a channel through which you can reach more users and promote your services. These new users have the potential to become your customers.|[Certification policy 1120.2](/legal/marketplace/certification-policies#11202-mobile-requirements)|
|Make your add-in commerce free on the iPad.|When it's running on the iPad, your add-in must be free of in-app purchases, trial offers, UI that aims to upsell to a non-free version, or links to any online stores where users can purchase or acquire other content, apps, or add-ins. Your Privacy Policy and Terms of Use pages must also be free of any commerce UI or AppSource links. Your add-in can still have commerce on other platforms. To do so, test the [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member) property and suppress all commerce when it returns `false`.|[Certification policy 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)|
|Submit your add-in to AppSource.|In Partner Center, on the **Product setup** page, select the **Make my product available on iOS and Android (if applicable)** check box, and provide your Apple developer ID in Account settings. Review the [Application Provider Agreement](https://go.microsoft.com/fwlink/?linkid=715691) to make sure you understand the terms.|[Make your solutions available in AppSource and within Office](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)|

### Detecting iPad devices

Your add-in can provide a different UI based on the device it's running on. To detect whether your add-in is running on an iPad, use the following APIs.

```javascript
const isTouchEnabled = Office.context.touchEnabled;
const allowCommerce = Office.context.commerceAllowed;

// On an iPad, touchEnabled returns true and commerceAllowed returns false
if (isTouchEnabled && !allowCommerce) {
    // Likely running on iPad - implement iPad-specific UI
    enableIPadInterface();
    hideCommerceFeatures();
}
```

### iPad development best practices

#### Develop and debug the add-in on Windows or Mac and sideload it to an iPad

You can't develop an add-in directly on an iPad, but you can develop and debug it on a Windows or Mac computer and sideload it to an iPad for testing. Since an add-in that runs in Office on iOS or Mac supports the same APIs as an add-in running in Office on Windows, your add-in's code should run the same way on these platforms. For details, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md) and [Sideload Office Add-ins on iPad for testing](../testing/sideload-an-office-add-in-on-ipad.md).

#### Specify API requirements in your add-in's manifest or with runtime checks

When you specify API requirements in your add-in's manifest, Office determines if the Office client app supports those API members. If the API members are available, your add-in will be available too.

Alternatively, you can perform a runtime check to determine if a method is available before using it in your add-in. Runtime checks ensure your add-in is always available and provides additional functionality when the methods are supported. For more information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).

### Outlook add-ins on iPad

For information about designing Outlook add-ins that look good and work well in Outlook on mobile devices, see [Add-ins for Outlook on mobile devices](../outlook/outlook-mobile-addins.md).

> [!NOTE]
> If you're using [Fluent UI React](../quickstarts/fluent-react-quickstart.md) for your design elements, many of these elements are built into the design system.

## Publish and market your add-in

Ready to share your add-in with the world? Here's how to get started.

- **Create a Partner Center account.** This process can take time, so if you plan to publish to AppSource, start early. See [Partner Center account](/partner-center/marketplace-offers/open-a-developer-account).

- **Create an effective AppSource listing.** Follow these tips:

  - Use succinct, descriptive titles (128 characters or fewer).
  - Write short, compelling descriptions that answer "What problem does this add-in solve?"
  - Convey your add-in's value proposition clearly in the title and description. Don't rely only on your brand.

  Learn more about [creating effective AppSource listings](/partner-center/marketplace-offers/create-effective-office-store-listings).

- **Publish to AppSource.** Follow the AppSource [prepublish checklist](/partner-center/marketplace-offers/checklist) and [submission guide](/partner-center/marketplace-offers/add-in-submission-guide). Make sure to:

  - Test your add-in thoroughly on all supported operating systems, browsers, and devices.
  - Provide detailed testing instructions and resources for certification reviewers.

- **Create a website.** Help users discover your add-in outside of AppSource.

- **Promote your add-in** from your website. See [how to promote your add-in](/partner-center/marketplace-offers/promote-your-office-store-solution).

> [!IMPORTANT]
> [!INCLUDE [AppSource enterprise info](../includes/appsource-enterprise.md)]

## Support older Microsoft webviews and Office versions (recommended but not required)

See [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
