---
title: Develop your Office Add-in to work with ITP when using third-party cookies
description: How to work with ITP and Office Add-ins when using third-party cookies.
ms.date: 09/24/2025
ms.localizationpriority: medium
---

# Develop your Office Add-in to work with ITP when using third-party cookies

If your Office Add-in needs third-party cookies for authentication or storing settings, you might run into issues with Intelligent Tracking Prevention (ITP). When the [Runtime](../testing/runtimes.md) that loads your add-in uses ITP, these cookies get blocked.

If your Office Add-in and website must rely on third-party cookies, follow these steps to work with ITP.

1. **Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749)** so that the authenticating domain (the third-party that expects cookies) forwards an authorization token to your website. Use the token to establish a first-party session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).
1. **Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/)** so that the third-party can request permission to access its first-party cookies. Current versions of Office on Mac and Office on the web both support this API.
    > [!NOTE]
    > For non-authentication purposes, consider using `localStorage` instead of cookies.
    >
    > However, starting in Version 115 of Chromium-based browsers like Chrome and Edge, [storage partitioning](https://developer.chrome.com/docs/privacy-sandbox/storage-partitioning/) is enabled to prevent cross-site tracking (see also [Microsoft Edge browser policies](/deployedge/microsoft-edge-policies#defaultthirdpartystoragepartitioningsetting)). This means storage APIs like local storage are only available to contexts with the same origin and top-level site.

The following code sample shows how to use the Storage Access API.

```javascript
function displayLoginButton() {
  const button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame.
      // Also you should have previously written a cookie when domain was loaded in the top frame.
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

Third-party cookies are cookies loaded in an iframe where the domain differs from the top-level frame. ITP can affect:

- **Complex authentication scenarios** where a pop-up dialog handles credentials and the add-in iframe needs cookie access to complete authentication.
- **Silent authentication scenarios** where you've previously authenticated through a pop-up but subsequent attempts try to authenticate through a hidden iframe.

### Platform-specific behavior

**Office on Mac**: Access to third-party cookies is blocked by the macOS Big Sur SDK. WKWebView ITP is enabled by default in Safari, and WKWebView blocks all third-party cookies. Office on Mac Version 16.44 (20121301) or later is integrated with the macOS Big Sur SDK.

**Safari browser**: End users can turn off ITP by toggling the **Prevent cross-site tracking** checkbox under **Preferences** > **Privacy**. Please note that ITP can't be disabled for the embedded WKWebView control in Office on Mac.

[!INCLUDE [chrome-tracking-prevention](../includes/chrome-tracking-prevention.md)]

## See also

- [Handle ITP in Safari and other browsers where third-party cookies are blocked](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Tracking Prevention in WebKit](https://webkit.org/tracking-prevention/)
- [Chrome’s “Privacy Sandbox”](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Introducing the Storage Access API](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)
