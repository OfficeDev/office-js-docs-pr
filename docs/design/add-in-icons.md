---
title: Icon guidelines for Office Add-ins
description: Get an overview of how to design icons and the Fresh and Monoline design styles for add-in commands.
ms.date: 11/03/2025
ms.topic: overview
ms.localizationpriority: medium
---

# Icons

Icons are the visual representation of a behavior or concept. They are often used to add meaning to controls and commands. Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment. They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.

> [!NOTE]
> This article about designing icons for ribbon buttons. For guidance about icons that represent the add-in in the app  acquisition and managment UIs of Microsoft 365 applications, see [Design icons for add-in acquisisiton and management](microsoft-365-extension-management-icons.md).

Office app ribbon interfaces have a standard visual style. This ensures consistency and familiarity across Office apps. The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.

Many HTML containers contain controls with iconography. Use Fabric Core’s custom font to render Office styled icons in your add-in. The icon font provided by [Fabric Core](fabric-core.md) contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.

## Design icons for add-in commands

[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI. Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command. The following articles provide stylistic and production guidelines to help you design icons that integrate seamlessly with Office.

- For the Monoline style of Microsoft 365, see [Monoline style icon guidelines for Office Add-ins](add-in-icons-monoline.md).
- For the Fresh style of perpetual Office 2016 and later, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).

> [!NOTE]
> You must choose one style or the other and your add-in will use the same icons whether it's running in Microsoft 365 or perpetual Office.

## See also

- [Add-in development best practices](../concepts/add-in-development-best-practices.md)
- [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md)
