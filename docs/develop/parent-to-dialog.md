---
title: Alternative ways of passing messages to a dialog box from its host page
description: Learn workarounds to use when the messageChild method isn't supported.
ms.date: 07/18/2022
ms.localizationpriority: medium
---

# Alternative ways of passing messages to a dialog box from its host page

The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), there are two other ways that you can pass information to the dialog box.

- Add query parameters to the URL that is passed to `displayDialogAsync`.
- Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*

> [!NOTE]
> \* There is a bug that will effect your strategy for token handling. If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.

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
> Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property). Again, *your code should neither read nor write to this value*.
