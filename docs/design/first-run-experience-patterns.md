---
title: First-run experience patterns for Office Add-ins
description: Learn the best practices for designing first-run experiences in Office Add-ins.
ms.date: 08/09/2024
ms.topic: best-practice
ms.localizationpriority: medium
---

# First-run experience patterns

A first-run experience (FRE) is a user's introduction to your add-in. An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in. This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in.

## Best practices

Follow these best practices when crafting your first-run experience.

|Do|Don't|
|:------|:------|
|Provide a simple and brief introduction to the main actions in the add-in. | Don't include information and call-outs that aren't relevant to getting started. |
|Give users the opportunity to complete an action that will positively impact their use of the add-in. | Don't expect users to learn everything at once. Focus on the action that provides the most value. |
|Create an engaging experience that users will want to complete. | Don't force the users to click through the first-run experience. Give users an option to bypass the first-run experience. |

Consider whether showing users the first-run experience once or periodically is important to your scenario. For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.

Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.

## Carousel

The carousel takes users through a series of features or informational pages before they start using the add-in.

*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*

![Illustration showing step 1 of a carousel in the first-run experience of an Office desktop application task pane. In this example, a "Skip" action is included in the top right of the task pane.](../images/add-in-FRE-step-1.png)

*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*

![Illustration showing step 2 of a carousel in the first-run experience of an Office desktop application task pane. In this example, there are 3 carousel screens in the task pane.](../images/add-in-FRE-step-2.png)

*Figure 3. Provide a clear call to action to exit the first-run experience*

![Illustration showing step 3 of a carousel in the first-run experience of an Office desktop application task pane. In this example, the third and final screen of the task pane shows a button to get started.](../images/add-in-FRE-step-3.png)

## Value placemat

The value placemat communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.

*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*

![Illustration showing a value placemat in the first-run experience of an Office desktop application task pane. In this example, the task pane displays the add-in logo, a description of the add-in, and a button to get started.](../images/add-in-FRE-value.png)

For an example that uses the value placemat pattern, see the [first-run experience tutorial](../tutorials/first-run-experience-tutorial.md).

### Video placemat

The video placemat shows users a video before they start using your add-in.

*Figure 5. First-run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*

![Illustration showing a video placemat in the first-run experience of an Office desktop application task pane.](../images/add-in-FRE-video.png)

*Figure 6. Video player - Users presented with a video within a dialog window*

![Illustration showing a video in a dialog window with an Office desktop application and add-in task pane in the background.](../images/add-in-FRE-video-dialog.png)

## See also

- [First-run experience tutorial](../tutorials/first-run-experience-tutorial.md)
- [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
