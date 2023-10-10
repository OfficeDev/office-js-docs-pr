---
title: Color guidelines for Office Add-ins
description: Learn how to use colors in the UI of an Office Add-in.
ms.date: 08/18/2023
ms.topic: best-practice
ms.localizationpriority: medium
---

# Color guidelines for Office Add-ins

Color is often used to emphasize brand and reinforce visual hierarchy. It helps identify an interface as well as guide customers through an experience. Inside Office, color is used for the same goals but it's applied purposefully and minimally. At no point does it overwhelm customer content. Even when each Office app is branded with its own dominant color, it's used sparingly.

![The color scheme for Office, Excel, Word, and PowerPoint. Major colors for Office are black and white, and minor colors are light gray, dark gray, and orange. The dominant color for Excel is green, Word is blue, and PowerPoint is orange.](../images/office-addins-color-schemes.png)

[Fabric Core](fabric-core.md) includes a set of default theme colors. When Fabric Core is applied to an Office Add-in in components or in layouts, the same goals apply. Color should communicate hierarchy, purposefully guiding customers to action without interfering with content. Fabric Core theme colors can introduce a new accent color to the overall interface. This new accent can conflict with Office app branding and interfere with hierarchy. In other words, Fabric Core can introduce a new accent color to the overall interface when used inside an add-in. This new accent color can distract and interfere with the overall hierarchy. Consider ways to avoid conflicts and interference. Use neutral accents or overwrite Fabric Core theme colors to match Office app branding or your own brand colors.

Office applications allow customers to personalize their interfaces by applying an Office UI theme. Customers can choose between four UI themes to vary styling of backgrounds and buttons in Word, PowerPoint, Excel, and other apps in the Office suite. To make your add-ins feel like a natural part of Office and respond to personalization, use our Theming APIs. For example, task pane background colors switch to a dark gray in some themes. Our theming APIs allow you to follow suit and adjust foreground text to ensure [accessibility](../design/accessibility-guidelines.md).

> [!NOTE]
>
> - For mail and task pane add-ins, use the [Context.officeTheme](/javascript/api/office/office.context) property to match the theme of the Office applications. This API is currently available in Office 2016 or later.
> - For PowerPoint content add-ins, see [Use Office themes in your PowerPoint add-ins](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Apply the following general guidelines for color.

- Use color sparingly to communicate hierarchy and reinforce brand.
- Overuse of a single accent color applied to both interactive and non-interactive elements can lead to confusion. For example, avoid using the same color for selected and unselected items in a navigation menu.
- Avoid unnecessary conflicts with Office branded app colors.
- Use your own brand colors to build association with your service or company.
- Ensure that all text is accessible. Be sure that there is a 4.5:1 contrast ratio between foreground text and background.
- Be aware of color blindness. Use more than just color to indicate interactivity and hierarchy.
- To learn more about designing add-in command icons with the Office icon color palette, see [icon guidelines](../design/add-in-icons.md).
