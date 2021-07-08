---
title: Develop your Office Add-in to work with ITP when using third-party cookies
description: How to work with ITP and Office Add-ins when using third-party cookies
ms.date: 03/12/2021
localization_priority: Normal
---

# Develop your Office Add-in to work with ITP when using third-party cookies

If your Office Add-in requires third-party cookies, those cookies are blocked if Intelligent Tracking Prevention (ITP) is used by the browser runtime that loaded your add-in. You may be using third-party cookies to authenticate users, or for other scenarios, such as storing settings.

If your Office Add-in and website must rely on third-party cookies, use the following steps to work with ITP:

1. Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) so that the authenticating domain (in your case, the third-party that expects cookies) forwards an authorization token to your website. Use the token to establish a first-party login session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).
2. Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) so that the third-party can request permission to get access to its first-party cookies. Current versions of Office on Mac and Office on the web both support this API.
    > [!NOTE]
    > If you're using cookies for purposes other than authentication, then consider using `localStorage` for your scenario.

The following code sample shows how to use the Storage Access API.

```javascript
function displayLoginButton() {
  var button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame
      // Also you should have previously written a cookie when domain was loaded in the top frame
      console.error("User cancelled or requirements were not met.");
    });
  });
}

if (document.hasStorageAccess) { 
  document.hasStorageAccess().then(function(hasStorageAccess) { 
    if (!hasStorageAccess) { 
      displayLoginButton(); 
    } else { 
      authenticateWithCookies(); 
    } 
  }); 
} else { 
    authenticateWithCookies(); 
} 
```

## About ITP and third-party cookies

Third-party cookies are cookies that are loaded in an iframe, where the domain is different from the top level frame. ITP could affect complex authentication scenarios, where a popup dialog is used to enter credentials and then the cookie access is needed by an add-in iframe to complete the authentication flow. ITP could also affect silent authentication scenarios, where you have previously used a popup dialog to authenticate, but subsequent use of the add-in tries to authenticate through a hidden iframe.

When developing Office Add-ins on Mac, access to third-party cookies is blocked by the MacOS Big Sur SDK. This is because WKWebView ITP is enabled by default on the Safari browser, and WKWebView blocks all third-party cookies. Office on Mac version 16.44 or later is integrated with the MacOS Big Sur SDK.

In the Safari browser, end users can toggle the **Prevent cross-site tracking** checkbox under **Preference** > **Privacy** to turn off ITP. However, ITP cannot be turned off for the embedded WKWebView control.

## See also

- [Handle ITP in Safari and other browsers where third-party cookies are blocked](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Tracking Prevention in WebKit](https://webkit.org/tracking-prevention/)
- [Chrome’s “Privacy Sandbox”](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Introducing the Storage Access API](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)