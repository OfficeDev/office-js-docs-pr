---
title: Window objects that are unsupported in Office Add-ins
description: This article specifies some of the window runtime objects that do not work in Office Add-ins.
ms.date: 05/20/2023
ms.localizationpriority: medium
---

# Window objects that are unsupported in Office Add-ins

For some versions of Windows and Office, add-ins run in a Trident (Internet Explorer 11) webview runtime. (For details, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Trident. These properties are disabled in add-ins to ensure that your add-in provides a consistent experience to all users, regardless of which browser or webview control the add-in is using. This also helps AngularJS load properly.

The following is a list of the disabled properties. The list is a work in progress. If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.

- `window.history.pushState`
- `window.history.replaceState`

## See also

- [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)