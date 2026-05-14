---
title: Pass data to a dialog box using local storage or query parameters
description: Use local storage or URL query parameters to pass data from a host page to a dialog box in your Office Add-in when the messageChild API isn't available.
ms.date: 05/14/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Pass data to a dialog box using local storage or query parameters

If your add-in runs on a platform or host that doesn't support the [DialogApi 1.2 requirement set](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), you can't use the [`messageChild`](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) method to send data from a host page to a dialog box. Instead, use one of the following approaches.

- **Local storage** - Write data to [`window.localStorage`](https://developer.mozilla.org/docs/Web/API/Window/localStorage) in the host page before opening the dialog. Both windows can access the same local storage if they share the same domain, including port number.
- **Query parameters** - Append key-value pairs to the URL you pass to [`displayDialogAsync`](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)). The dialog reads the values when it opens.

For the recommended approach using `messageChild`, see [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box).

## Choose an approach

| Consideration | Local storage | Query parameters |
| --- | --- | --- |
| **Data size** | Suitable for larger payloads. | Best for small values. |
| **Data availability** | Available anytime after the dialog opens. | Available in the URL when the dialog loads.<br><br>**Tip**: Save the values to variables to read them while the dialog remains open. |
| **Persistence across navigation** | Persists until explicitly cleared. | Lost if the dialog redirects. |

## Use local storage

Call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as shown in the following example.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Code in the dialog box reads the item when it's needed, as shown in the following example.

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

[!INCLUDE [browser-security-updates](../includes/browser-security-updates.md)]

## Use query parameters

Append key-value pairs to the URL you pass to `displayDialogAsync`. This approach works best for small values that the dialog needs only when it first opens.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Code in your dialog box can parse the URL and read the parameter value.

> [!IMPORTANT]
> Office automatically adds a query parameter called `_host_info` to the URL passed to `displayDialogAsync`. It appends this parameter after your custom query parameters, if any. It doesn't append `_host_info` to any subsequent URLs that the dialog box navigates to. Microsoft might change the content of this value, or remove it entirely, in the future, so your code shouldn't read it. Office adds the same value to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property). Again, *your code should neither read nor write to this value*.

## Troubleshoot common issues

- If local storage appears empty in the dialog, verify the host page and dialog use the exact same origin.
- If a query parameter value is missing, confirm it's present in the initial URL passed to `displayDialogAsync`.
- If data is needed after redirects in the dialog flow, use local storage or server state instead of relying only on query parameters.

## See also

- [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md)
- [Best practices and rules for the Office Dialog API](dialog-best-practices.md)
- [Handle errors and events in the Office dialog box](dialog-handle-errors-events.md)
