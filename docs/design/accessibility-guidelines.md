---
title: Accessibility guidelines for Office Add-ins
description: Learn how to make your Office Add-in accessible to all users, including keyboard navigation, screen reader support, high contrast, and WCAG compliance.
ms.date: 04/15/2026
ms.topic: best-practice
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Accessibility guidelines for Office Add-ins

Accessible add-ins work for everyone &mdash; people who use screen readers, navigate by keyboard, rely on high contrast, or interact through touch. Building with accessibility in mind from the start improves usability for all your users and expands the market for your solutions.

The Web Content Accessibility Guidelines (WCAG) are the international standards that define what's needed for your add-in to be accessible. We recommend you become familiar with them before you start building.

- [Web Content Accessibility Guidelines (WCAG) 2.2](https://www.w3.org/TR/WCAG22/)
- [Web Content Accessibility Guidelines (WCAG) 3.0 (Preview)](https://www.w3.org/TR/WCAG22/)
- [WCAG standards and resources](/compliance/regulatory/offering-wcag-2-1)
- [W3C WAI tutorials](https://www.w3.org/WAI/tutorials/)
- [Web accessibility principles and guidelines](/training/modules/web-accessibility-principles-guidelines/) training module on Microsoft Learn (17 minute overview)


## Design for multiple input methods

Your add-in should support keyboard, touch, and mouse input so that every user can interact with it regardless of their device or abilities.

- Ensure that users can complete all operations by using [keyboard interactions](/windows/apps/design/input/keyboard-interactions) alone. Users should be able to reach every actionable element on the page by using a combination of <kbd>Tab</kbd> and the arrow keys.
- On mobile devices, your add-in should provide useful audio feedback when users operate controls by touch.
- Preserve a logical reading and navigation order in the DOM or UI tree.
- Ensure all interactive elements expose their name, role, and state to assistive technologies by using appropriate ARIA labels.

## Make your add-in easy to use

Predictable behavior and clear feedback help all users understand what's happening in your add-in.

- Don't rely on sound alone to alert users to important information.
- Don't rely on color, shape, size, or visual location alone to convey meaning or instructions.
- Don't require complex gestures (drag, multi‑touch, timed motion) without providing simpler alternatives.
- [Manage focus](/windows/apps/develop/input/focus-navigation) carefully. Don’t move focus to a different element unless the user initiates the change.
- Provide a way to verify, confirm, or reverse all binding actions.
- Don't impose a time limit for user action.

## Make your add-in easy to see

Clear visual design and sufficient contrast ensure that users with low vision or color blindness can read and understand your UI.

- Avoid unexpected color changes.
- Provide meaningful, timely descriptions for UI elements, titles, headings, inputs, and errors. Ensure that control names clearly describe the intent of the control.
- Verify that your UI elements render correctly in Windows [high contrast themes](/windows/apps/design/accessibility/high-contrast-themes).
- Follow [WCAG color contrast guidelines](https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum). Aim for a contrast ratio of at least 4.5:1 for normal text and 3:1 for large text.

## Content and media

- Provide text alternatives for all meaningful non‑text content (for example, images, icons, SVGs, charts, and custom controls). 
- Don't put meaningful text inside images unless an equivalent text alternative is provided.
- Provide transcripts for audio‑only content.
- Provide synchronized captions for prerecorded or live video with audio.
- Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.
- Provide a way for users to pause, stop, or mute audio or animation.

## Test for accessibility

Regular [accessibility testing](/windows/apps/design/accessibility/accessibility-testing) helps you catch issues before they reach users and ensures your add-in continues to work with assistive technologies as standards and platforms evolve.

### Integrate automated checks

Automated tools can detect many issues, but they don’t evaluate how well your add-in works with assistive technologies.

- Run [Accessibility Insights](https://accessibilityinsights.io/) as part of your build or CI pipeline to identify common accessibility issues early.
- Use automated checks to validate:
  - Color contrast.
  - Accessible names and labels.
  - Keyboard focus order.
  - ARIA roles and attributes.
  - Target size and spacing.
 
### Validate assistive technology support

Test your add-in with commonly used assistive technologies to confirm that users can:

- Navigate all interactive elements using a keyboard alone.
- Understand the purpose of controls through screen reader output.
- Complete tasks without relying on vision, hearing, or precise pointer movements.

Verify the experience using:

- [Windows Narrator](https://support.microsoft.com/windows/e4397a0d-ef4f-b386-d8ae-c172f109bdb1)
- [JAWS](https://support.freedomscientific.com/Downloads/JAWS)
- [NVDA](https://www.nvaccess.org/download/).

Confirm that:

- All functionality is available without using a mouse.
- Focus indicators are visible and not obscured.
- Dynamic updates are announced appropriately.
- Custom UI components expose their role, state, and value.

### Test interaction patterns

Evaluate user flows &mdash; not just individual controls &mdash; to ensure that:

- Dragging or gesture-based interactions have a keyboard-accessible alternative.
- Time-limited interactions can be paused, extended, or disabled.
- Alerts and notifications are conveyed without relying on sound alone.
- Authentication and input methods do not depend on memory, vision, or timed responses.

### Re-test as standards evolve

Assistive technologies and international accessibility guidelines change over time. Test periodically to keep up with changes to the international accessibility guidelines. For more information, see [Accessibility testing](/windows/apps/design/accessibility/accessibility-testing).

## See also

- [Overview of accessibility in Windows apps](/windows/apps/design/accessibility/accessibility-overview)
- [WebView2 accessibility APIs](/microsoft-edge/webview2/concepts/overview-features-apis)
- [Accessibility in the Store](/windows/apps/design/accessibility/accessibility-in-the-store)
- [Developing for Web Accessibility](https://www.w3.org/WAI/tips/developing/)
- [Accessibility Fundamentals learning path](/training/paths/accessibility-fundamental/)
- [European Accessibility Act (EAA)](https://www.deque.com/blog/european-accessibility-act-eaa-top-20-key-questions-answered/)
