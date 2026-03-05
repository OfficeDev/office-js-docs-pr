---
title: Address same-origin policy limitations in Office Add-ins
description: Learn how to accommodate same-origin policy limitations with JSONP, CORS, iframes, and other techniques.
ms.date: 03/06/2026
ms.localizationpriority: medium
---

# Address same-origin policy limitations in Office Add-ins

The browser or webview control enforces the same-origin policy, which prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This policy means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy prevents a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it's hosted.

Because Office Add-ins are hosted in a webview control, the same-origin policy applies to script running in their web pages as well.

The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains. You can use a few common techniques to securely overcome same-origin policy enforcement. This article provides only the briefest introduction to some of them. Use the links provided to get started in your research of these techniques.

## Implement server-side code using a token-based authorization scheme

One way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.

## Use cross-origin resource sharing (CORS)

To learn more about cross-origin resource sharing, see the many resources available on the web, such as [Cross-Origin Resource Sharing (CORS)](https://web.dev/cross-origin-resource-sharing/).

> [!NOTE]
> If your add-in uses event-based activation or integrated spam reporting in Outlook, you must also configure a well-known URI to enable CORS or SSO. For details, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Office Add-in](use-sso-in-event-based-activation.md).

## Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)

For an example of how to build your own proxy by using IFRAME and POST MESSAGE, see [Cross-Window Messaging](https://johnresig.com/blog/cross-window-messaging/).

## Use JSONP for anonymous access

Another way to overcome same-origin policy limitations is to use [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.

The following example of JSONP uses a technique that works in any Office Add-in.

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    const script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}
```

> [!NOTE]
> JSONP is a legacy technique has higher security risk than modern alternatives. Use JSONP only if you must maintain older integrations that can't be migrated yet. For new development, prefer CORS or a server-side proxy.

## See also

- [Same-origin policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)
- [Cross-Origin Resource Sharing (CORS)](https://developer.mozilla.org/docs/Web/HTTP/CORS)
- [Authenticate and authorize with the Office dialog API](auth-with-office-dialog-api.md)
- [Authorization with non-Microsoft identity providers](auth-external-add-ins.md)
- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
