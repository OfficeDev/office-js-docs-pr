---
title: Accessibility guidelines for Office Add-ins
description: Learn how to make your Office Add-in accessible to all users.
ms.date: 12/3/2024
ms.topic: best-practice
ms.localizationpriority: medium
---

# Accessibility guidelines

As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Engineering and implementing inclusive experiences provide better usability and customer satisfaction, as well as a larger market for your solutions. We recommend you become familiar with the Web Content Accessibility Guidelines (WCAG), international web standards that define what's needed for your add-in to be accessible.

- [Explore the WCAG standards and resources](/compliance/regulatory/offering-wcag-2-1)
- [Explore the WCAG tutorials](https://www.w3.org/WAI/tutorials/)

Apply the following guidelines to ensure that your solution is accessible to all audiences.

## Design for multiple input methods

- Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the <kbd>Tab</kbd> and arrow keys.
- On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.
- Provide helpful labels for all interactive controls.
- [Explore more design and UI resources.](/windows/apps/design/accessibility/accessibility)

## Make your add-in easy to use

- Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.
- Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.
- Provide a way to verify, confirm, or reverse all binding actions.
- Provide a way to pause or stop media, such as audio and video.
- Don't impose a time limit for user action.

## Make your add-in easy to see

- Avoid unexpected color changes.
- Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.
- Verify you UI elements render correctly in the Windows high-contrast themes.
- Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.

## Account for assistive technologies

- Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.
- Don't provide text in an image format. Screen readers can't read text within images.
- Provide a way for users to adjust or mute all audio sources.
- Provide a way for users to turn on captions or audio description with audio sources.
- Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.

## Test your add-in

- Always use accessibility verification and testing tools like [Accessibility Insights](https://accessibilityinsights.io/) on your add-in to catch and resolve issues before you ship.
- Verify the screen reading experience using [Windows Narrator](https://support.microsoft.com/windows/e4397a0d-ef4f-b386-d8ae-c172f109bdb1), [JAWS](https://support.freedomscientific.com/Downloads/JAWS), or [NVDA](https://www.nvaccess.org/download/).
- Periodically run the tools to keep up with changes to the international accessibility guidelines. For more information, see [Accessibility testing](/windows/apps/design/accessibility/accessibility-testing).

## See also

- [Accessibility in the Store](/windows/apps/design/accessibility/accessibility-in-the-store)
- [Web Content Accessibility Guidelines (WCAG) 2.2](https://www.w3.org/TR/WCAG22/)
- [Developing for Web Accessibility](https://www.w3.org/WAI/tips/developing/)
- [Accessibility Fundamentals Learning Path](/training/paths/accessibility-fundamental/)
- [European Accessibility Act (EAA)](https://www.deque.com/blog/european-accessibility-act-eaa-top-20-key-questions-answered/)
