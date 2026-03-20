---
title: UX design patterns for Office Add-ins
description: Get an overview of the UI design patterns for Office Add-ins, including patterns for navigation, authentication, first-run, and branding.
ms.date: 10/29/2025
ms.topic: overview
ms.localizationpriority: medium
---

# UX design patterns for Office Add-ins

Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.  

Our UX patterns are composed of components. Components are controls that help your customers interact with elements of your software or service. Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.

[Fluent UI React components](../quickstarts/fluent-react-quickstart.md) look and behave like a part of Office, as do the framework-neutral components of [Office UI Fabric JS](fabric-core.md). Take advantage of either set of components to integrate with Office. Alternatively, if your add-in has its own preexisting component language, you don't need to discard it. Look for opportunities to retain it while integrating with Office. Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.

The provided patterns are best practice solutions based on common customer scenarios and user experience research. They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own. Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.

Use the UX pattern templates to:

- Apply solutions to common customer scenarios.
- Apply design best practices.
- Incorporate [Fluent UI](add-in-design.md) components and styles.
- Build add-ins that visually integrate with the default Office UI.
- Ideate and visualize UX.

## Getting started

The patterns are organized by key actions or experiences that are common in an add-in. The main groups are:

- [First-run experience (FRE)](../design/first-run-experience-patterns.md)
- [Authentication](../design/authentication-patterns.md)
- [Navigation](../design/navigation-patterns.md)
- [Branding Design](../design/branding-patterns.md)

Browse each grouping to get an idea of how you can design your add-in using best practices.

> [!NOTE]
> The example screens shown throughout the design pattern documentation are designed and displayed at a resolution of **1366x768**.

## See also

- [Office Add-in design language](add-in-design-language.md)
- [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
- [Fluent UI React in Office Add-ins](../quickstarts/fluent-react-quickstart.md)
