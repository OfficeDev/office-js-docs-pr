---
title: Alternative ways of passing messages to a dialog box from its host page
description: Learn workarounds to use when the messageChild method isn't supported.
ms.date: 03/06/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Alternative ways of passing messages to a dialog box from its host page

The recommended way to pass data and messages from a parent page to a child dialog is by using the `messageChild` method, as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in runs on a platform or host that doesn't support the [DialogApi 1.2 requirement set](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), you can use two other ways to pass information to the dialog.

- Store the information somewhere accessible to both the host window and dialog. The two windows don't share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).

  [!INCLUDE [browser-security-updates](../includes/browser-security-updates.md)]

- Add query parameters to the URL that is passed to `displayDialogAsync`.

## Use local storage

To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as shown in the following example.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Code in the dialog box reads the item when it's needed, as shown in the following example.

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

## Use query parameters

Use this approach for small values that are needed only when the dialog initially opens.

In the host page, add query parameters to the URL passed to `displayDialogAsync`.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Code in your dialog box can parse the URL and read the parameter value.

> [!IMPORTANT]
> Office automatically adds a query parameter named `_host_info` to the URL that it passes to `displayDialogAsync`. It appends this parameter after your custom query parameters, if any. It doesn't append `_host_info` to any subsequent URLs that the dialog box navigates to. Microsoft might change the content of this value, or remove it entirely, in the future, so your code shouldn't read it. Office adds the same value to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property). Again, *your code should neither read nor write to this value*.

## Troubleshoot common issues

- If local storage appears empty in the dialog, verify the host page and dialog use the exact same origin.
- If a query parameter value is missing, confirm it is present in the initial URL passed to `displayDialogAsync`.
- If data is needed after redirects in the dialog flow, use local storage or server state instead of relying only on query parameters.

## See also

- [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md)
- [Best practices and rules for the Office Dialog API](dialog-best-practices.md)
- [Handle errors and events in the Office dialog box](dialog-handle-errors-events.md)
