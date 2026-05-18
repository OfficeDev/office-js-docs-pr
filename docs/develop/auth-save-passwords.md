---
title: Automatically save passwords in Office Add-ins with WebView2
description: Configure an Office Add-in running in WebView2 to prompt users to save credentials after sign-in.
ms.topic: how-to
ms.localizationpriority: medium
ms.date: 05/18/2026
---

# Enable automatic password saving in Microsoft Edge WebView2

If your add-in prompts users to sign in, Microsoft Edge WebView2 can offer to save credentials and fill them the next time users sign in. This article shows the minimum HTML and JavaScript pattern to trigger the WebView2 save-password prompt in Office Add-ins on Windows.

## Sign-in flow

First, create a sign-in form in your add-in's HTML page. The form should include username and password fields, and a button to submit the credentials. Use the standard HTML credential fields: `type="text"` for the username and `type="password"` for the password. The example is the minimum HTML you need.

```html
<div>
  <label for="username">Username:</label>
  <input type="text" id="username" name="username" />

  <label for="password">Password:</label>
  <input type="password" id="password" name="password" />

  <button id="btn" type="button">Sign in</button>
</div>
```

In the sign-in button's click event handler, call an authentication library to sign in the user. After sign-in completes, redirect to a new page. When WebView2 detects the redirect and credential fields, it can prompt users to save credentials.

```javascript
async function btnSignIn() {
  // Get the username and password credentials entered by the user.
  const username = document.getElementById("username").value;
  const pwd = document.getElementById("password").value;

  try {
    // Sign in the user. This is a placeholder for the actual sign-in logic.
    await signInUser(username, pwd);

    // Redirect to a success page to trigger the password autosave.
    window.location.href = "/home.html";
  }
  catch (error) {
    console.error("Sign in failed: " + error);
    return;
  }
}
```

## What users see

When users enter credentials in your add-in and the add-in redirects to a new page, WebView2 asks whether they want to save the username and password. The next time your add-in prompts for credentials, WebView2 can fill in the saved account information.

:::image type="content" source="../images/edge-webview2-automatic-save-passwords.png" alt-text="The dialog from WebView2 prompting the user if they want to save their username and password.":::

## Remove saved credentials

Users can remove saved passwords by deleting the WebView2 local cache folder at `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2\`. If your add-in relies on automatic password saving, document this folder location so users can remove saved credentials.

## Related content

- [Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2)
- [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)
- [Authorization with non-Microsoft identity providers](auth-external-add-ins.md)
- [Enable single sign-on in an Office add-in with nested app authentication](enable-nested-app-authentication-in-your-add-in.md)
