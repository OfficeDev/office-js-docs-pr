---
title: Design the UI of Office Add-ins
description: Learn the best practices for the visual design of Office Add-ins.
ms.date: 10/18/2022
ms.topic: best-practice
ms.localizationpriority: high
---


# Design the UI of Office Add-ins

Office Add-ins extend the Office experience by providing contextual functionality that users can access within Office clients. Add-ins empower users to get more done by enabling them to access external functionality within Office, without costly context switches.

Your add-in UI design must integrate seamlessly with Office to provide an efficient, natural interaction for your users. Take advantage of [add-in commands](add-in-commands.md) to provide access to your add-in and apply the best practices that we recommend when you create a custom HTML-based UI.

## Office design principles

Office applications follow a general set of interaction guidelines. The applications share content and have elements that look and behave similarly. This commonality is built on a set of design principles. The principles help the Office team create interfaces that support customers’ tasks. Understanding and following them will help you support your customers’ goals inside of Office.

Follow the Office design principles to create positive add-in experiences.

- **Design explicitly for Office.** The functionality, as well as the look and feel, of an add-in must harmoniously complement the Office experience. Add-ins should feel native. They should fit seamlessly into Word on an iPad or PowerPoint on the web. A well-designed add-in will be an appropriate blend of your experience, the platform, and the Office application. Apply document and UI theming where appropriate. Consider using Fluent UI for the web as your design language and tool set. The Fluent UI for the web has two flavors.

  - **For non-React UIs:** Use **Fabric Core**, an open-source collection of CSS classes and Sass mixins that give you access to colors, animations, fonts, icons, and grids. (It's called "Fabric Core" instead of "Fluent Core" for historical reasons.) To get started, see [Fabric Core in Office Add-ins](fabric-core.md).
  
  [!INCLUDE [alert-fluent-ui-web-components](../includes/alert-fluent-ui-web-components.md)]

  - **For React UIs:** use **Fluent UI React**, a React front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products. It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS. To get started, see [Fluent UI React in Office Add-ins](../quickstarts/fluent-react-quickstart.md).

- **Favor content over chrome.** Allow customers’ page, slide, or spreadsheet to remain the focus of the experience. An add-in is an auxiliary interface. No accessory chrome should interfere with the add-in’s content and functionality. Brand your experience wisely. We know it's important to provide users with a unique, recognizable experience but avoid distraction. Strive to keep the focus on content and task completion, not brand attention.

- **Make it enjoyable and keep users in control.** People enjoy using products that are both functional and visually appealing. Craft your experience carefully. Get the details right by considering every interaction and visual detail. Allow users to control their experience. The necessary steps to complete a task must be clear and relevant. Important decisions should be easy to understand. Actions should be easily reversible. An add-in is not a destination – it’s an enhancement to Office functionality.

- **Design for all platforms and input methods**. Add-ins are designed to work on all the platforms that Office supports, and your add-in UX should be optimized to work across platforms and form factors. Support mouse/keyboard and touch input devices, and ensure that your custom HTML UI is responsive to adapt to different form factors. For more information, see [Touch](../concepts/add-in-development-best-practices.md#optimize-for-touch).

## See also

- [Add-in development best practices](../concepts/add-in-development-best-practices.md)
