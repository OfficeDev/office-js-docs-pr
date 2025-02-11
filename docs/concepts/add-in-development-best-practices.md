---
title: Best practices for developing Office Add-ins
description: Apply the best practices when developing Office Add-ins.
ms.topic: best-practice
ms.date: 10/10/2024
ms.localizationpriority: medium
---

# Best practices for developing Office Add-ins

Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## Provide clear value

- Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:
  - Make core authoring tasks faster and easier, with fewer interruptions.
  - Enable new scenarios within Office.
  - Embed complementary services within Office applications.
  - Improve the Office experience to enhance productivity.
- Make sure that the value of your add-in is clear to users right away by [creating an engaging first-run experience](#create-an-engaging-first-run-experience).
- Learn how to create an [effective AppSource listing](/partner-center/marketplace-offers/create-effective-office-store-listings) to promote your add-in.
  - Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.
  - Your add-in must provide sufficient value to justify users' investment. Your add-in shouldn't be a simple utility or of limited scope.
  - [!INCLUDE [AppSource enterprise info](../includes/appsource-enterprise.md)]

## Create an engaging first-run experience

- Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.

- Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.

- Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.

- Provide teaching UI to guide users and make your UI personal.

  ![A "Do" versus "Don't" comparison on how to guide your users to use the UI. The "Do" example shows an add-in that includes a button users can click to get started. The "Don't" example shows an add-in with no introductory steps or buttons.](../images/contoso-part-catalog-do-dont.png)

- If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.

  ![A "Do" versus "Don't" comparison on including an option to insert sample data in your add-in. The "Do" example shows an add-in that includes a button users can click to insert sample data. The "Don't" example shows an add-in without sample data or buttons.](../images/add-in-title.png)

- Offer [free trials](/partner-center/marketplace-offers/monetize-addins-through-microsoft-commercial-marketplace). If your add-in requires a subscription, make some functionality available without a subscription.

- Make the sign-up experience simple. Prefill information, such as email and display name, and skip email verifications.

- Avoid pop-up windows. If you have to use them, guide the user to enable your pop-up window.

For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).

## Use add-in commands

- Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).

## Apply UX design principles

- Ensure that the look and feel and functionality of your add-in complements the Office experience. See [Design the UI of Office Add-ins](../design/add-in-design.md).

- Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.

- Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.

- Use branding to inspire trust and orient users. Don't use branding to overwhelm or advertise to users.

- Avoid scrolling. Optimize for 1366 x 768 resolution.

- Don't include unlicensed images.

- Use [clear and simple language](../design/voice-guidelines.md) in your add-in.

- Account for [accessibility](../design/accessibility-guidelines.md) - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.

- Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.

### Optimize for touch

- Use the [Context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member) property to detect whether the Office application that your add-in runs on is touch enabled.

  > [!NOTE]
  > This property isn't supported in Outlook.

- Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.

- Don't rely on non-touch input methods like hover or right-click.

- Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.

- Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad.md).

> [!NOTE]
> If you're using [Fluent UI React](../quickstarts/fluent-react-quickstart.md) for your design elements, many of these elements are built into the design system.

## Optimize and monitor add-in performance

- Create the perception of fast UI responses. Your add-in should load in 500 ms or less.

- Ensure that all user interactions respond in under one second.

- Provide loading indicators for long-running operations.

- Use a content delivery network (CDN) to host images, resources, and common libraries. Load as much as you can from one place.

- Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.

- If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).

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

- Monitor your service health, and use telemetry to monitor user success.

- Minimize data exchanges between the add-in and the Office document. For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).

## Publish and market your add-in

- Create a [Partner Center account](/partner-center/marketplace-offers/open-a-developer-account). This could take some time. If you plan to publish your add-in to AppSource, get this process going as soon as possible.

- Create an [effective AppSource listing](/partner-center/marketplace-offers/create-effective-office-store-listings).

  - Use succinct and descriptive add-in titles. Include no more than 128 characters.

  - Write short, compelling, and accurate descriptions of your add-in. Answer the question "What problem does this add-in solve?".

  - Convey the value proposition of your add-in in your title and description. Don't rely on your brand.

- Create a website to help users find and use your add-in.

- Publish your add-in to [AppSource](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center). Be sure to follow the AppSource [prepublish checklist](/partner-center/marketplace-offers/checklist) and [submission guide](/partner-center/marketplace-offers/add-in-submission-guide).

  - Ensure that your add-in works error-free in all operating systems, browsers, and devices supported by the minimum set of APIs declared in your add-in manifest.

  - Provide detailed testing instructions and resources for certification reviewers to verify the functionality of your add-in.

- [Promote your add-in](/partner-center/marketplace-offers/promote-your-office-store-solution) from your website.

> [!IMPORTANT]
> [!INCLUDE [AppSource enterprise info](../includes/appsource-enterprise.md)]

## Support older Microsoft webviews and Office versions (recommended but not required)

See [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
