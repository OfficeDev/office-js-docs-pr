---
title: Design Office Add-ins
description: ''
ms.date: 06/20/2019
localization_priority: Normal
---


# Design Office Add-ins

Office Add-ins extend the Office experience by providing contextual functionality that users can access within Office clients. Add-ins empower users to get more done by enabling them to access third-party functionality within Office, without costly context switches. 

Your add-in UX design must integrate seamlessly with Office to provide an efficient, natural interaction for your users. Take advantage of [add-in commands](add-in-commands.md) to provide access to your add-in and apply the best practices that we recommend when you create custom HTML-based UI.

## Office design principles

Office applications follow a general set of interaction guidelines. The apps share content and have elements that look and behave similarly. This commonality is built on a set of design principles. The principles help the Office team create interfaces that support customers’ tasks. Understanding and following them will help you support your customers’ goals inside of Office.

Follow the Office design principles to create positive add-in experiences:

- **Design explicitly for Office.** The functionality, look and feel of an add-in must harmoniously complement the Office experience. Add-ins should feel native. They should fit seamlessly into Word on an iPad or PowerPoint on the web. A well-designed add-in will be an appropriate blend of your experience, the platform and the Office application. Consider using Office UI Fabric as your design language. Apply document and UI theming where appropriate.

- **Focus on a few key tasks; do them well.** Help customers get one job done without getting in the way of other jobs. Provide real value to customers. Focus on common use cases, pick carefully those that benefit users most when interacting with Office documents.

- **Favor content over chrome.** Allow customers’ page, slide or spreadsheet to remain the focus of the experience. An add-in is an auxiliary interface. No accessory chrome should interfere with the add-in’s content and functionality. Brand your experience wisely. We know it is important to provide users with a unique, recognizable experience but avoid distraction. Strive to keep the focus on content and task completion, not brand attention.

- **Make it enjoyable and keep users in control.** People enjoy using products that are both functional and visually appealing. Craft your experience carefully. Get the details right by considering every interaction and visual detail. Allow users to control their experience. The necessary steps to complete a task must be clear and relevant. Important decisions should be easy to understand. Actions should be easily reversible. An add-in is not a destination – it’s an enhancement to Office functionality.

- **Design for all platforms and input methods**. Add-ins are designed to work on all the platforms that Office supports, and your add-in UX should be optimized to work across platforms and form factors. Support mouse/keyboard and touch input devices, and ensure that your custom HTML UI is responsive to adapt to different form factors. For more information, see [Touch](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## See also
- [Office UI Fabric](https://developer.microsoft.com/fabric) 
- [Add-in development best practices](../concepts/add-in-development-best-practices.md)

