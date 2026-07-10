---
title: Authorization with non-Microsoft identity providers
description: Learn how to use OAuth 2.0 flows to authorize an Office Add-in to access non-Microsoft services and data sources.
ms.date: 07/10/2026
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Authorization with non-Microsoft identity providers

Many identity providers, in addition to the Microsoft identity platform, can work with your add-in. These providers let users grant an Office Add-in access to their accounts in other services.

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

A fundamental idea of OAuth is that an application can be a [security principal](/windows/security/identity-protection/access-control/security-principals) unto itself, just like a user or a group, with its own identity and set of permissions. In a typical flow, a user takes an action in the add-in that requires another service. The add-in requests a specific set of permissions to that user's account. The service then prompts the user to grant those permissions.

After permission is granted, the service sends the add-in an encoded *access token*. The add-in includes the token in requests to the service's APIs. The token grants only the permissions that the user approved, and it expires after a specified time.

## Choose an OAuth 2.0 flow

Several OAuth patterns, called *flows* or *grant types*, are designed for different scenarios. The following two patterns are the most commonly implemented.

- **Implicit flow**: Communication between the add-in and the online service is implemented with client-side JavaScript. This flow is commonly used in single-page applications (SPAs).
- **Authorization Code flow**: Communication is *server-to-server* between your add-in's web application and the online service. So, it is implemented with server-side code.

The purpose of an OAuth flow is to secure the identity and authorization of the application. In the Authorization Code flow, the identity provider issues a *client secret* that must remain confidential. An application that has no server-side back end, such as an SPA, can't safely store that secret, so we recommend the Implicit flow for SPAs.

You should be familiar with the pros and cons of the Implicit flow and the Authorization Code flow. For more information about these two flows, see [Authorization Code](https://tools.ietf.org/html/rfc6749#section-1.3.1) and [Implicit](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## Use the Implicit flow in Office Add-ins

Check the identity provider's documentation to confirm that it supports the Implicit flow.

For information about libraries that support the Implicit flow, see the **Libraries** section later in this article.

## Use the Authorization Code flow in Office Add-ins

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For some examples, see the **Libraries** section later in this article.

## Libraries

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

- **Facebook**: Search [Facebook for Developers](https://developers.facebook.com) for "library" or "sdk".
- **General OAuth 2.0**: The IETF OAuth Working Group maintains [OAuth Code](https://oauth.net/code/), a page of library links for more than a dozen languages. Some of these libraries are for implementing an OAuth-compliant service. For an Office Add-in, look for *client* libraries, because your web server is a client of the OAuth-compliant service.

## Middleman services

Your add-in can use a middleman service such as [OAuth.io](https://oauth.io) or [Auth0](https://auth0.com) to perform authorization. A middleman service might provide access tokens for popular online services, simplify social sign-in for your add-in, or both. Your add-in can connect to the middleman service with either client-side script or server-side code, and the middleman service returns any required tokens for the online service.

We recommend that the UI for authentication and authorization in your add-in use the Office dialog API to open a sign-in page. For more information, see [Authenticate and authorize with the Office dialog API](auth-with-office-dialog-api.md).

When you open an Office dialog this way, the dialog runs in a separate browser and JavaScript-engine instance from the parent page, such as the add-in's task pane or function file. A token, and any other information that can be converted to a string, is passed back to the parent by using `messageParent`. The parent page can then use the token to make authorized calls to the resource.

Because of this architecture, be careful when you use APIs from a middleman service. Some services provide an API set in which the code creates a context object that both gets a token and uses that token in later calls to the resource. Some services even use a single API method that both makes the initial call and creates the context object. An object like this can't be fully stringified, so it can't be passed from the Office dialog to the parent page.

Middleman services typically also provide a second API set at a lower level of abstraction, such as a REST API. This lower-level API set usually includes one API that gets a token from the service and other APIs that pass the token back to the service when requesting access to the resource. Use this lower-level API set so the Office dialog can get the token and then pass it to the parent page by using `messageParent`.

## What is CORS?

CORS stands for [Cross-Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about using CORS in add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).

## See also

- [Overview of authentication and authorization in Office Add-ins](overview-authn-authz.md)