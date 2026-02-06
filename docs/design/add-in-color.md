---
title: Color guidelines for Office Add-ins
description: Learn how to use colors in the UI of an Office Add-in.
ms.date: 10/28/2025
ms.topic: best-practice
ms.localizationpriority: medium
---

# Color guidelines for Office Add-ins

Color is often used to emphasize brand and reinforce visual hierarchy. It helps identify an interface as well as guide customers through an experience. Inside Office, color is used for the same goals but it's applied purposefully and minimally. At no point does it overwhelm customer content. Even when each Office app is branded with its own dominant color, it's used sparingly.

:::image type="content" source="../images/office-addins-color-schemes.png" alt-text="The color scheme for Office, Excel, Word, and PowerPoint. Major colors for Office are black and white, and minor colors are light gray, dark gray, and orange. The dominant color for Excel is green, Word is blue, and PowerPoint is orange.":::

## Use UI component libraries
[Fluent UI React](../quickstarts/fluent-react-quickstart.md) and [Fabric Core](fabric-core.md) include a set of default theme colors. When these libraries are applied to the components or layouts of an Office Add-in, the same goals apply. Color should communicate hierarchy, purposefully guiding customers to action without interfering with content. Theme colors from Fluent UI React or Fabric Core can introduce a new accent color to the overall interface. These accent colors can conflict with Office app branding and the hierarchy. Consider ways to avoid conflicts and interference. Use neutral accents or overwrite theme colors to match Office app branding or your own brand colors.

## Develop with theming APIs
In Office applications, customers personalize their interfaces by applying an [Office UI theme](https://support.microsoft.com/office/365-63e65e1c-08d4-4dea-820e-335f54672310). Customers choose between four UI themes to vary styling of backgrounds and buttons in Excel, Outlook, PowerPoint, Word, and other apps in the Office suite. To make your add-ins feel like a natural part of Office and respond to personalization, use our [Theming APIs](/javascript/api/office/office.officetheme). For example, task pane background colors switch to a dark gray in some themes. With the Theming APIs, follow suit and adjust foreground text to ensure [accessibility](../design/accessibility-guidelines.md).

> [!TIP]
>
> - To ensure that your add-in applies the correct color combinations in its UI, test it with all four Office themes, including the **Use system setting** option.
> - For guidance on how to dynamically match the design of your PowerPoint add-in with the current document or Office theme, see [Use document themes in your PowerPoint add-ins](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

## Apply best practices

- Use color sparingly to communicate hierarchy and reinforce brand.
- Overuse of a single accent color applied to both interactive and non-interactive elements can lead to confusion. For example, avoid using the same color for selected and unselected items in a navigation menu.
- Avoid unnecessary conflicts with Office branded app colors.
- Use your own brand colors to build association with your service or company.
- Ensure that all text is [accessible](accessibility-guidelines.md). Be sure that there is a 4.5:1 contrast ratio between foreground text and background.
- Be aware of color blindness. Use more than just color to indicate interactivity and hierarchy.
- To learn more about designing add-in command icons with the Office icon color palette, see [icon guidelines](../design/add-in-icons.md).

## See also

- [Office Add-in design language](add-in-design-language.md)
- [Accessibility guidelines](accessibility-guidelines.md)
- [Branding patterns](branding-patterns.md)
- [Fluent UI React in Office Add-ins](../quickstarts/fluent-react-quickstart.md)
