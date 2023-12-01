---
title: Alternative ways of passing messages to a dialog box from its host page
description: Learn workarounds to use when the messageChild method isn't supported.
ms.date: 11/16/2023
ms.topic: how-to
ms.localizationpriority: medium
---

# Alternative ways of passing messages to a dialog box from its host page

The recommended way to pass data and messages from a parent page to a child dialog is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that doesn't support the [DialogApi 1.2 requirement set](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), there are two other ways that you can pass information to the dialog.

- Store the information somewhere accessible to both the host window and dialog. The two windows don't share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).

  [!INCLUDE [browser-security-updates](../includes/browser-security-updates.md)]

- Add query parameters to the URL that is passed to `displayDialogAsync`.

## Use local storage

To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Code in the dialog box reads the item when it's needed, as in the following example.

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

## Use query parameters

The following example shows how to pass data with a query parameter.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Code in your dialog box can parse the URL and read the parameter value.

> [!IMPORTANT]
> Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It isn't appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code shouldn't read it. The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property). Again, *your code should neither read nor write to this value*.
