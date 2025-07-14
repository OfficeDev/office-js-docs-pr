---
title: Addressing same-origin policy limitations in Office Add-ins
description: Learn how to accommodate same-origin policy limitations with JSONP, CORS, iframes, and other techniques.
ms.date: 12/21/2023
ms.localizationpriority: medium
---

# Addressing same-origin policy limitations in Office Add-ins

The same-origin policy enforced by the browser or webview control prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.

Because Office Add-ins are hosted in a webview control, the same-origin policy applies to script running in their web pages as well.

The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains. There are a few common techniques for securely overcoming same-origin policy enforcement. This article can only provide the briefest introduction to some of them. Please use the links provided to get started in your research of these techniques.

## Use JSONP for anonymous access

One way to overcome same-origin policy limitations is to use [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.

The following is an example of JSONP that uses a technique that will work in any Office Add-in.

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

## Implement server-side code using a token-based authorization scheme

Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.

## Use cross-origin resource sharing (CORS)

To learn more about cross-origin resource sharing, see the many resources available on the web, such as [Cross-Origin Resource Sharing (CORS)](https://web.dev/cross-origin-resource-sharing/).

> [!NOTE]
> For information on how to use CORS in an Outlook add-in that implements event-based activation or integrated spam reporting, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Outlook add-in](../develop/use-sso-in-event-based-activation.md).

## Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)

For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](https://johnresig.com/blog/cross-window-messaging/).

## See also

- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
