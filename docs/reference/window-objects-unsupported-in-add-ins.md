---
title: Window objects that are unsupported in Office Add-ins
description: This article specifies some of the window runtime objects that do not work in Office add-ins.
ms.date: 07/10/2020
localization_priority: Normal
---

# Window objects that are unsupported in Office Add-ins

For some versions of Windows and Office, add-ins run in an Internet Explorer 11 runtime. (For details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Internet Explorer 11. To ensure that your add-in provides a consistent experience to all users regardless of which browser the add-in is using, these properties and subproperties are disabled in add-ins. Disabling these properties also helps AngularJS load properly in add-ins.

The following is a partial list of the disabled properties. The list is a work in progress. If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.

- `window.history.pushState`
- `window.history.replaceState`
