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
> The [Office Add-in design principles](../design/add-in-design.md) apply to Outlook mobile add-ins. Review those guidelines in addition to the mobile-specific patterns outlined in this article.

## Add-in components and patterns

A typical mobile add-in is made up of the following components.

- Branding area
- Navigation bar
- Section title
- Cells or input fields
- Actions

The following images show how each component appears in Outlook on Android and on iOS.

**Android**
:::image type="content" source="../images/outlook-mobile-design-overview-android.jpg" alt-text="Diagram of basic UX patterns for a task pane on Android.":::

**iOS**
:::image type="content" source="../images/outlook-mobile-design-overview.png" alt-text="Diagram of basic UX patterns for a task pane on iOS.":::

The succeeding sections outline UX design patterns for Outlook on mobile. To learn more about design patterns for Office Add-ins, see [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md).

### Loading page

When a user taps your add-in, display the UI as quickly as possible. If there's any delay, use a progress bar or activity indicator. Use a progress bar when the duration is known and an activity indicator when the duration is unknown.

**An example of loading pages on Android**
:::image type="content" source="../images/outlook-mobile-design-loading-android.jpg" alt-text="Examples of a progress bar and an activity indicator on Android.":::

**An example of loading pages on iOS**
:::image type="content" source="../images/outlook-mobile-design-loading.png" alt-text="Examples of a progress bar and an activity indicator on iOS.":::

### Sign-in or sign-up page

Make your sign-in and sign-up flows straightforward and easy to use.

**An example of a sign-in page on Android**
:::image type="content" source="../images/outlook-mobile-design-signin-android.png" alt-text="Example of a sign-in page on Android.":::

**An example of sign-in and sign-up pages on iOS**
:::image type="content" source="../images/outlook-mobile-design-signin.png" alt-text="Example of sign-in and sign-up pages on iOS.":::

### Brand bar

Include your branding element on the first screen of your add-in. The brand bar provides recognition and helps set context for the user.

**Do**:

- Use a white version of your logo and set the background banner to your main brand color.
- Stay within prescribed [margins](#margins) and 60% coverage.

**Don't**:

- Go outside margins.
- Repeat the brand bar on subsequent pages. The navigation bar already contains the name of your company or your brand.

**An example of branding on Android**
:::image type="content" source="../images/outlook-mobile-design-branding-android.png" alt-text="Examples of brand bars on Android.":::

**An example of branding on iOS**
:::image type="content" source="../images/outlook-mobile-design-branding.png" alt-text="Examples of brand bars on iOS.":::

> [!NOTE]
> Ads shouldn't be shown within add-ins in Outlook on iOS or on Android.

### Margins

The recommended mobile margins vary depending on the Outlook client.

- **Android**: 16px for each side
- **iOS**: 15px (8% of screen) for each side

The following is an example of margins set in Outlook on iOS.

:::image type="content" source="../images/outlook-mobile-design-margins.png" alt-text="Examples of margins on iOS.":::

### Typography

Use simple typography to make content easy to scan. The following table outlines the typography guidelines for Outlook on Android and on iOS.

| Component | Android | iOS |
| --- | --- | --- |
| Title 1 | Medium 20pt | Light 28pt |
| Title 2 | Semibold 24pt | Regular 22pt |
| Subheader | Medium 14pt | Regular 15pt |
| Body 1 | Regular 16pt | Regular 17pt |
| Body 2 | Regular 14pt | Not applicable |
| Caption | Regular 12pt | Regular 12pt |
| Footnote | Not applicable | Regular 13pt |
| Button | Medium 14pt | Medium 14pt |

**Typography on Android**
:::image type="content" source="../images/outlook-mobile-design-typography-android.png" alt-text="Typography samples for Android.":::

**Typography on iOS**
:::image type="content" source="../images/outlook-mobile-design-typography.png" alt-text="Typography samples for iOS.":::

### Color palette

Use color subtly in Outlook on mobile. Limit color usage to actions and error states and use a unique color only for the brand bar.

The following table outlines the recommended color palette for add-in components on Outlook mobile. Color codes are shown in hexadecimal format.

| Color | Component |
| --- | --- |
| #222222 | Most headlining types |
| #545454 | Secondary headlining type |
| #8E8E93 | Icons only |
| #999999 | Secondary icon type |
| #E1E1E1 | Dividers or disabled states |
| #F8F8F8 | Form background or section title bar |
| #0075DA | Actions |
| #E63237 | Error states or damaging actions |

:::image type="content" source="../images/outlook-mobile-design-color-palette.png" alt-text="Color palette for Outlook on mobile devices.":::

### Cells

Cells, such as section titles and input fields, display the content of your add-in.

#### Section title

**Do**:

- Use section titles to label pages since the navigation bar can't be used to label a page.
- Use a colored section title to drive attention and maintain content hierarchy.
- Use a colored section at the top of the page to display account information or settings.
- (iOS only) Use the tall version of a section title if the title appears between two selectable elements. The space that the tall version offers ensures that the correct element is selected.

**Don't**:

- Apply center alignment to a section title.
- (Android only) Add a section title above tabs.
- (iOS only) Use the small version of a section title if the title appears between two selectable elements. This increases the likelihood of selecting the incorrect element.
- (iOS only) Use the small version of a section title if it appears above tabs.

**Examples of section titles on Android**
:::image type="content" source="../images/outlook-mobile-design-cell-type-android.png" alt-text="Section title examples for Android.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-cell-dos-android.png" alt-text="Section title 'dos' for Android.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-cell-donts-android.png" alt-text="Section title 'don'ts' for Android.":::

**Examples of small and tall section titles on iOS**
:::image type="content" source="../images/outlook-mobile-design-cell-types.png" alt-text="Small and tall section title examples for iOS.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-cell-dos.png" alt-text="Section title 'dos' for iOS.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-cell-donts.png" alt-text="Section title 'don'ts' for iOS.":::

#### Input fields

Use the appropriate input field type for your content. For example, use radio buttons when users should select only one option.

**Do**:

- Be mindful of the number of items when stacking multiple input fields. Group items together as needed and use a section title to identify the group when appropriate.
- (iOS only) Align divider lines with the left edge of the text when stacking multiple items. Don't align them with the item's icon on the left.

**Examples of input fields on Android**
:::image type="content" source="../images/outlook-mobile-design-cell-input-1-android.png" alt-text="Input fields on Android, including radio buttons and checkboxes.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-cell-input-2-android.png" alt-text="Text input fields on Android.":::

**Example of input fields on iOS**
:::image type="content" source="../images/outlook-mobile-design-cell-input.png" alt-text="Checked options on iOS.":::

### Actions

Even if your add-in handles many actions, focus on the most important ones.

**Do**:

- Stack a maximum of two actions.
- Anchor actions to the bottom of the page when the content length is unknown or extends beyond one page.
- Float actions when content fits on a single page.
- (iOS only) Implement a **Back** action to help with navigation. Anchor the **Back** action to the top of the page.

**Examples of actions on Android**
:::image type="content" source="../images/outlook-mobile-design-action-cells-android.png" alt-text="Actions and cells in Android.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-action-dos-android.png" alt-text="Actions 'dos' for Android.":::

**Examples of actions on iOS**
:::image type="content" source="../images/outlook-mobile-design-action-cells.png" alt-text="Actions and cells in iOS.":::
* * *
:::image type="content" source="../images/outlook-mobile-design-action-dos.png" alt-text="Actions 'dos' for iOS.":::

### Buttons

Use buttons when other UI elements appear below them. For example, use a button on a sign-in page when links to service terms and privacy policies appear below it. In contrast, use actions when they're the last element on the page, anchored to the bottom. 

**Do**:

- Use the appropriate button style to indicate button state. For example, use a disabled state (dimmed button) when an option is unavailable.
- Float a button if there are other UI elements, such as text, displayed below it.

**Examples of buttons on Android**
:::image type="content" source="../images/outlook-mobile-design-buttons-android.png" alt-text="Examples of buttons for Android.":::

**Examples of buttons on iOS**
:::image type="content" source="../images/outlook-mobile-design-buttons.png" alt-text="Examples of buttons for iOS.":::

### Tabs

Use tabs to organize your add-in's content.

**Examples of tabs on Android**
:::image type="content" source="../images/outlook-mobile-design-tabs-android.png" alt-text="Examples of tabs for Android.":::

**Examples of tabs on iOS**
:::image type="content" source="../images/outlook-mobile-design-tabs.png" alt-text="Examples of tabs for iOS.":::

### Icons

When possible, follow the current Outlook mobile design for icons. Use the standard size and color.

**Examples of icons on Android**
:::image type="content" source="../images/outlook-mobile-design-icons-android.jpg" alt-text="Examples of icons for Android.":::

**Examples of icons on iOS**
:::image type="content" source="../images/outlook-mobile-design-icons.png" alt-text="Examples of icons for iOS.":::

## Add-in examples

When Outlook mobile add-ins launched, we worked closely with partners building add-ins. To showcase the potential of their add-ins on Outlook mobile, our designer created end-to-end flows for each add-in using these guidelines and patterns.

> [!IMPORTANT]
> These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.

### GIPHY

**An example of GIPHY on Android**
:::image type="content" source="../images/outlook-mobile-design-giphy-android.png" alt-text="End-to-end design for the GIPHY add-in on Android.":::

**An example of GIPHY on iOS**
:::image type="content" source="../images/outlook-mobile-design-giphy.png" alt-text="End-to-end design for the GIPHY add-in on iOS.":::

### Nimble

**An example of Nimble on Android**
:::image type="content" source="../images/outlook-mobile-design-nimble-android.png" alt-text="End-to-end design for the Nimble add-in on Android.":::

**An example of Nimble on iOS**
:::image type="content" source="../images/outlook-mobile-design-nimble.png" alt-text="End-to-end design for the Nimble add-in on iOS.":::

### Dynamics CRM

**An example of Dynamics CRM on Android**
:::image type="content" source="../images/outlook-mobile-design-crm-android.png" alt-text="End-to-end design for the Dynamics CRM add-in on Android.":::

**An example of Dynamics CRM on iOS**
:::image type="content" source="../images/outlook-mobile-design-crm.png" alt-text="End-to-end design for the Dynamics CRM add-in on iOS.":::

## See also

- [Design the UI of Office Add-ins](../design/add-in-design.md)
- [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md)
