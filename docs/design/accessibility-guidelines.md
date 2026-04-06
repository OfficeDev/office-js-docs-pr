---
title: Accessibility guidelines for Office Add-ins
description: Learn how to make your Office Add-in accessible to all users, including keyboard navigation, screen reader support, high contrast, and WCAG compliance.
ms.date: 04/06/2026
ms.topic: best-practice
ms.localizationpriority: medium
---

# Accessibility guidelines for Office Add-ins

Accessible add-ins work for everyone — people who use screen readers, navigate by keyboard, rely on high contrast, or interact through touch. Building with accessibility in mind from the start improves usability for all your users and expands the market for your solutions.

The Web Content Accessibility Guidelines (WCAG) are the international standards that define what's needed for your add-in to be accessible. We recommend you become familiar with them before you start building.

- [Web Content Accessibility Guidelines (WCAG) 2.2](https://www.w3.org/TR/WCAG22/)
- [WCAG standards and resources](/compliance/regulatory/offering-wcag-2-1)
- [W3C WAI tutorials](https://www.w3.org/WAI/tutorials/)

## Design for multiple input methods

Your add-in should support keyboard, touch, and mouse input so that every user can interact with it regardless of their device or abilities.

- Ensure that users can complete all operations by using only the keyboard. Users should be able to reach every actionable element on the page by using a combination of <kbd>Tab</kbd> and the arrow keys.
- On mobile devices, provide useful audio feedback when users operate controls by touch.
- Provide helpful labels for all interactive controls.

For more design and UI resources, see [Accessibility in Windows apps](/windows/apps/design/accessibility/accessibility).

## Make your add-in easy to use

Predictable behavior and clear feedback help all users understand what's happening in your add-in.

- Don't rely on a single attribute — such as color, size, shape, location, orientation, or sound — to convey meaning in your UI.
- Avoid unexpected changes of context, such as moving focus to a different UI element without user action.
- Provide a way to verify, confirm, or reverse all binding actions.
- Provide a way to pause or stop media, such as audio and video.
- Don't impose a time limit for user action.

## Make your add-in easy to see

Clear visual design and sufficient contrast ensure that users with low vision or color blindness can read and understand your UI.

- Avoid unexpected color changes.
- Provide meaningful, timely descriptions for UI elements, titles, headings, inputs, and errors. Ensure that control names clearly describe the intent of the control.
- Verify that your UI elements render correctly in Windows high contrast themes.
- Follow [WCAG color contrast guidelines](https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum). Aim for a contrast ratio of at least 4.5:1 for normal text and 3:1 for large text.

## Account for assistive technologies

Screen readers, magnifiers, and other assistive tools rely on well-structured content and semantic markup to present your add-in to users.

- Avoid features that interfere with assistive technologies, including visual, audio, or other interactions.
- Don't put text in images. Screen readers can't read text within images.
- Provide a way for users to adjust or mute all audio sources.
- Provide a way for users to turn on captions or audio descriptions.
- Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.

## Test for accessibility

Regular testing catches issues before they reach your users and helps you stay current with evolving standards.

- Run [Accessibility Insights](https://accessibilityinsights.io/) on every build of your add-in to catch issues before you ship.
- Verify the screen reader experience with [Windows Narrator](https://support.microsoft.com/windows/e4397a0d-ef4f-b386-d8ae-c172f109bdb1), [JAWS](https://support.freedomscientific.com/Downloads/JAWS), or [NVDA](https://www.nvaccess.org/download/).
- Test periodically to keep up with changes to the international accessibility guidelines. For more information, see [Accessibility testing](/windows/apps/design/accessibility/accessibility-testing).

## See also

- [Design the UI of Office Add-ins](add-in-design.md)
- [Color guidelines for Office Add-ins](add-in-color.md)
- [Custom keyboard shortcuts in Office Add-ins](keyboard-shortcuts.md)
- [Accessibility in the Store](/windows/apps/design/accessibility/accessibility-in-the-store)
- [Web Content Accessibility Guidelines (WCAG) 2.2](https://www.w3.org/TR/WCAG22/)
- [Developing for Web Accessibility](https://www.w3.org/WAI/tips/developing/)
- [Accessibility Fundamentals learning path](/training/paths/accessibility-fundamental/)
- [European Accessibility Act (EAA)](https://www.deque.com/blog/european-accessibility-act-eaa-top-20-key-questions-answered/)
