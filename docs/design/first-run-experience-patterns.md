---
title: First-run experience patterns for Office Add-ins
description: Learn the best practices for designing first-run experiences in Office Add-ins.
ms.date: 08/12/2025
ms.topic: best-practice
ms.localizationpriority: medium
---

# First-run experience patterns

A well-designed first-run experience (FRE) introduces users to your Office Add-in. It guides them through its core features and benefits. An effective FRE not only helps users get started quickly but also increases engagement and retention by making a positive first impression.

## Why first-run experiences matter

The initial moments with your add-in set the tone for future interactions. A clear, concise FRE can:

- Highlight your add-in's unique value.
- Reduce user confusion and support requests.
- Encourage users to take meaningful actions.
- Foster ongoing usage and loyalty.

## Key principles for first-run experiences

- **Keep it simple**: Focus on the essential actions and benefits. Avoid overwhelming users with too much information.
- **Be actionable**: Provide clear next steps or calls to action that help users realize immediate value.
- **Offer flexibility**: Allow users to skip or revisit the FRE as needed.
- **Engage visually**: Use graphics, carousels, or videos to make the experience memorable and easy to follow.

## FRE patterns for Office Add-ins

Explore these proven patterns to design an effective first-run experience:

- [Carousel](#carousel): Guide users through a sequence of screens highlighting features and benefits, with options to advance or skip.
- [Value placemat](#value-placemat): Present your add-in’s value proposition, logo, feature summary, and a call-to-action in a single, visually engaging layout.
- [Video placemat](#video-placemat): Use a short video to demonstrate your add-in’s capabilities, paired with a clear call-to-action.

Choose the pattern that best fits your add-in's complexity and user needs. Consider showing the FRE periodically if your add-in is used infrequently, to help users stay familiar with its features.

## Best practices

Follow these best practices when crafting your first-run experience.

| Do | Don't |
|:------|:------|
| Provide a simple and brief introduction to the main actions in the add-in. | Don't include information and call-outs that aren't relevant to getting started. |
| Give users the opportunity to complete an action that will positively impact their use of the add-in. | Don't expect users to learn everything at once. Focus on the action that provides the most value. |
| Create an engaging experience that users will want to complete. | Don't force the users to click through the first-run experience. Give users an option to bypass the first-run experience. |

Consider whether showing users the first-run experience once or periodically is important to your scenario. For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.

Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.

## Carousel

The carousel takes users through a series of features or informational pages before they start using the add-in.

*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*

:::image type="content" source="../images/add-in-FRE-step-1.png" alt-text="Illustration showing step 1 of a carousel in the first-run experience of an Office desktop application task pane. In this example, a 'Skip' action is included in the top right of the task pane.":::

*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*

:::image type="content" source="../images/add-in-FRE-step-2.png" alt-text="Illustration showing step 2 of a carousel in the first-run experience of an Office desktop application task pane. In this example, there are 3 carousel screens in the task pane.":::

*Figure 3. Provide a clear call to action to exit the first-run experience*

:::image type="content" source="../images/add-in-FRE-step-3.png" alt-text="Illustration showing step 3 of a carousel in the first-run experience of an Office desktop application task pane. In this example, the third and final screen of the task pane shows a button to get started.":::

## Value placemat

The value placemat communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.

*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*

:::image type="content" source="../images/add-in-FRE-value.png" alt-text="Illustration showing a value placemat in the first-run experience of an Office desktop application task pane. In this example, the task pane displays the add-in logo, a description of the add-in, and a button to get started.":::

For an example that uses the value placemat pattern, see the [first-run experience tutorial](../tutorials/first-run-experience-tutorial.md).

## Video placemat

The video placemat shows users a video before they start using your add-in.

*Figure 5. First-run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*

:::image type="content" source="../images/add-in-FRE-video.png" alt-text="Illustration showing a video placemat in the first-run experience of an Office desktop application task pane.":::

*Figure 6. Video player - Users presented with a video within a dialog window*

:::image type="content" source="../images/add-in-FRE-video-dialog.png" alt-text="Illustration showing a video in a dialog window with an Office desktop application and add-in task pane in the background.":::

## See also

- [First-run experience tutorial](../tutorials/first-run-experience-tutorial.md)
- [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
