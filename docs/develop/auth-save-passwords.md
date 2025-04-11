---
title: Automatically save passwords in Microsoft Edge WebView2
description: Learn how to enable WebView2 to save passwords when users sign in using your add-in.
ms.localizationpriority: medium
ms.date: 03/10/2025
---

# Enable automatic password saving in Microsoft Edge WebView2

Most browsers can automatically save passwords on behalf of the user when they sign in. This helps users manage passwords in a secure environment. Microsoft Edge WebView2 also supports automatic password saving. When your add-in is loaded in Microsoft Office on Windows, Webview2 hosts your add-in. To enable automatic password saving, add HTML input controls for the username and password, as shown in the following HTML.

```html
<div>
    <label for="username">Username:</label><br/>
    <input type="text" id="username" name="username" /><br/>
    
    <label for="password">Password:</label><br/>
    <input type="password" id="password" name="password" /><br/>
    
    <button id="btn" type="button">Sign in</button>
</div>
```

In the button click event handler for the sign-in button, call the authentication library of your choice to sign in the user. Once the sign-in is complete, redirect to a new web page. When WebView2 sees the redirect, and the username and password, it prompts the user to offer to automatically save the credentials. The following code shows how to handle the sign-in button click event.

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

## How the user manages passwords

When the user enters a new password in your add-in, and your add-in redirects to a new web page, WebView2 asks the user if they want to save their username and password. The next time your add-in prompts for credentials, WebView2 automatically enters the user's account information.

:::image type="content" source="../images/edge-webview2-automatic-save-passwords.png" alt-text="The dialog from WebView2 prompting the user if they want to save their username and password.":::

Users remove saved passwords by deleting The WebView2 local cache folder at `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2\`. If your add-in relies on automatically saving passwords, you should document this folder location so users can remove their passwords.

## Related content

- [Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2)
- [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)
