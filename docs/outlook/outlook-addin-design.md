---
title: Design add-ins for Outlook on mobile devices
description: Guidelines to help you design and build a compelling add-in for Outlook on Android and on iOS.
ms.date: 12/02/2025
ms.topic: best-practice
ms.localizationpriority: high
---

# Design add-ins for Outlook on mobile devices

Outlook on mobile devices provides a unique environment for add-ins, with platform-specific design patterns for Android and iOS. This article provides guidelines and visual examples to help you create add-ins that feel native to each mobile platform while maintaining a consistent brand experience.

> [!TIP]
> The [general Office Add-in design principles](../design/add-in-design.md) apply to Outlook mobile add-ins. Review those guidelines in addition to the mobile-specific patterns outlined in this article.

## Add-in components and patterns

A typical mobile add-in is made up of the following components.

- Branding area
- Navigation bar
- Section title
- Cells or input fields
- Actions

The following images show how each component appears in Outlook on Android and on iOS.

**Android**
![Diagram of basic UX patterns for a task pane on Android.](../images/outlook-mobile-design-overview-android.jpg)

**iOS**
![Diagram of basic UX patterns for a task pane on iOS.](../images/outlook-mobile-design-overview.png)

The succeeding sections outline UX design patterns for Outlook on mobile. To learn more about design patterns for Office Add-ins, see [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md).

### Loading

When a user taps on your add-in, the UX should display as quickly as possible. If there is any delay, use a progress bar or activity indicator. A progress bar should be used when the amount of time is determinable and an activity indicator should be used when the amount of time is indeterminable.

**An example of loading pages on Android**

![Examples of a progress bar and an activity indicator on Android.](../images/outlook-mobile-design-loading-android.jpg)

**An example of loading pages on iOS**

![Examples of a progress bar and an activity indicator on iOS.](../images/outlook-mobile-design-loading.png)

### Sign in/Sign up

Make your sign in and sign up flows straightforward and simple to use.

**An example sign in page on Android**

![Examples of page to sign in on Android.](../images/outlook-mobile-design-signin-android.png)

**An example page to sign in and sign up on iOS**

![Examples of pages to sign in and sign up on iOS.](../images/outlook-mobile-design-signin.png)

### Brand bar

The first screen of your add-in should include your branding element. Designed for recognition, the brand bar also helps set context for the user. Because the navigation bar contains the name of your company or your brand, it's unnecessary to repeat the brand bar on subsequent pages.

**An example of branding on Android**

![Examples of brand bars on Android.](../images/outlook-mobile-design-branding-android.png)

**An example of branding on iOS**

![Examples of brand bars on iOS.](../images/outlook-mobile-design-branding.png)

> [!NOTE]
> Ads should not be shown within add-ins in Outlook on iOS or on Android.

### Margins

The recommended mobile margins vary depending on the Outlook client.

- **Android**: 16px for each side
- **iOS**: 15px (8% of screen) for each side

The following is an example of margins set in Outlook on iOS.

![Examples of margins on iOS.](../images/outlook-mobile-design-margins.png)

### Typography

Keep typography simple to make content easy to scan.

**Typography on Android**

![Typography samples for Android.](../images/outlook-mobile-design-typography-android.png)

**Typography on iOS**

![Typography samples for iOS.](../images/outlook-mobile-design-typography.png)

### Color palette

Color usage is subtle in Outlook on mobile. We recommend localizing color usage to actions and error states, with only the brand bar using a unique color.

![Color palette for Outlook on mobile devices.](../images/outlook-mobile-design-color-palette.png)

### Cells

Since the navigation bar can't be used to label a page, use section titles to label pages.

**Examples of cells on Android**

![Cell types for Android.](../images/outlook-mobile-design-cell-type-android.png)
* * *
![Cell 'dos' for Android.](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![Cell 'don'ts' for Android.](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![Cells and inputs for Android part 1.](../images/outlook-mobile-design-cell-input-1-android.png)

![Cells and inputs for Android part 2.](../images/outlook-mobile-design-cell-input-2-android.png)

**Examples of cells on iOS**

![Cell types for iOS.](../images/outlook-mobile-design-cell-types.png)
* * *
![Cell 'dos' for iOS.](../images/outlook-mobile-design-cell-dos.png)
* * *
![Cell 'don'ts' for iOS.](../images/outlook-mobile-design-cell-donts.png)
* * *
![Cells and inputs for iOS.](../images/outlook-mobile-design-cell-input.png)

### Actions

Even if your app handles a multitude of actions, think about the most important ones you want your add-in to perform, and concentrate on those.

**Examples of actions on Android**

![Actions and cells in Android.](../images/outlook-mobile-design-action-cells-android.png)
* * *
![Actions 'dos' for Android.](../images/outlook-mobile-design-action-dos-android.png)

**Examples of actions on iOS**

![Actions and cells in iOS.](../images/outlook-mobile-design-action-cells.png)
* * *
![Actions 'dos' for iOS.](../images/outlook-mobile-design-action-dos.png)

### Buttons

Buttons are used when there are other UX elements below (vs. actions, where the action is the last element on the screen).

**Examples of buttons on Android**

![Examples of buttons for Android.](../images/outlook-mobile-design-buttons-android.png)

**Examples of buttons on iOS**

![Examples of buttons for iOS.](../images/outlook-mobile-design-buttons.png)

### Tabs

Tabs can aid in content organization.

**Examples of tabs on Android**

![Examples of tabs for Android.](../images/outlook-mobile-design-tabs-android.png)

**Examples of tabs on iOS**

![Examples of tabs for iOS.](../images/outlook-mobile-design-tabs.png)

### Icons

Icons should follow the current Outlook mobile design when possible. Use our standard size and color.

**Examples of icons on Android**

![Examples of icons for Android.](../images/outlook-mobile-design-icons-android.jpg)

**Examples of icons on iOS**

![Examples of icons for iOS.](../images/outlook-mobile-design-icons.png)

## Add-in examples

> [!NOTE]
> [!INCLUDE [Calendar add-ins not available in Teams](../includes/calendar-availability.md)]

When Outlook mobile add-ins were launched, we worked closely with our partners who were building add-ins. As a way to showcase the potential of their add-ins on Outlook mobile, our designer put together end-to-end flows for each add-in, leveraging our guidelines and patterns.

> [!IMPORTANT]
> These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.

### GIPHY

**An example of GIPHY on Android**

![End-to-end design for the GIPHY add-in on Android.](../images/outlook-mobile-design-giphy-android.png)

**An example of GIPHY on iOS**

![End-to-end design for the GIPHY add-in on iOS.](../images/outlook-mobile-design-giphy.png)

### Nimble

**An example of Nimble on Android**

![End-to-end design for the Nimble add-in on Android.](../images/outlook-mobile-design-nimble-android.png)

**An example of Nimble on iOS**

![End-to-end design for the Nimble add-in on iOS.](../images/outlook-mobile-design-nimble.png)

### Dynamics CRM

**An example of Dynamics CRM on Android**

![End-to-end design for the Dynamics CRM add-in on Android.](../images/outlook-mobile-design-crm-android.png)

**An example of Dynamics CRM on iOS**

![End-to-end design for the Dynamics CRM add-in on iOS.](../images/outlook-mobile-design-crm.png)

## See also

- [Design the UI of Office Add-ins](../design/add-in-design.md)
- [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md)
