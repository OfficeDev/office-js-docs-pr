---
title: Authorize external services in your Office Add-in
description: ''
ms.date: 12/04/2017
localization_priority: Priority
---

# Authorize external services in your Office Add-in

Popular online services, including Office 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in.

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

A fundamental idea of OAuth is that an application can be a security principal unto itself, just like a user or a group, with its own identity and set of permissions. In the most typical scenarios, when the user takes an action in the Office Add-in that requires the online service, the add-in sends the service a request for a specific set of permissions to the user's account. The service then prompts the user to grant the add-in those permissions. After the permissions are granted, the service sends the add-in a small encoded *access token*. The add-in can use the service by including the token in all its requests to the service's APIs. But the add-in can act only within the permissions that the user granted it. The token also expires after a specified time.

Several OAuth patterns, called *flows* or *grant types*, are designed for different scenarios. The following two patterns are the most commonly implemented:

- **Implicit flow**: Communication between the add-in and the online service is implemented with client-side JavaScript.
- **Authorization Code flow**: Communication is *server-to-server* between your add-in's web application and the online service. So, it is implemented with server-side code.

The purpose of an OAuth flow is to secure the identity and authorization of the application. In the Authorization Code flow, you're provided a *client secret* that needs to be kept hidden. A Single Page Application (SPA) has no way to protect the secret, so we recommend that you use the Implicit flow in SPAs.

You should be familiar with the pros and cons of the Implicit flow and the Authorization Code flow. For more infomation about these two flows, see [Authorization Code](https://tools.ietf.org/html/rfc6749#section-1.3.1) and [Implicit](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## Using the Implicit flow in Office Add-ins
The best way to find out if an online service supports the Implicit flow is to consult the service's documentation. For services that support the Implicit flow, you can use the **Office-js-helpers** JavaScript library to do all the detailed work for you:

- [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

For information about other libraries that support the Implicit flow, see the **Libraries** section later in this article.

## Using the Authorization Code flow in Office Add-ins

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For more information about some of these libraries, see the **Libraries** section later in this article.

The following samples provide examples of add-ins that implement the Authorization Code flow:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

### Relay/Proxy functions

You can use the Authorization Code flow even with a serverless web application by storing the **client ID** and **client secret** values in a simple function that is hosted in a service such as [Azure Functions](https://azure.microsoft.com/services/functions) or [Amazon Lambda](https://aws.amazon.com/lambda).
The function exchanges a given code for an **access token** and relays it back to the client. The security of this approach depends on how well access to the function is guarded.

To use this technique, your add-in displays a UI/popup to show the login screen for the online service (Google, Facebook, and so on). When the user signs in and grants the add-in permission to her resources in the online service, the add-in receives a code which can be then sent to the online function. The services described in the **Middleman services** section later in this article use a similar flow.

## Libraries

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

**Office 365 and other services that use Azure Active Directory as the authorization provider**: [Azure Active Directory Authentication Libraries](https://azure.microsoft.com/documentation/articles/active-directory-authentication-libraries/). A preview is also available for the [Microsoft Authentication Library](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

**Facebook**: Search [Facebook for Developers](https://developers.facebook.com) for "library" or "sdk".

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](https://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## Middleman services

Your add-in can use a middleman service such as OAuth.io or Auth0 to perform authorization. A middleman service may either provide access tokens for popular online services or simplify the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman service and it will send your add-in any required tokens for the online service. All of the authorization implementation code is in the middleman service.

For examples of add-ins that use a middleman service for authorization, see the following samples:

- [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0) uses Auth0 to enable social login with Facebook, Google, and Microsoft Accounts.

- [Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io) uses OAuth.io to get access tokens from Facebook and Google.

## What is CORS?

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about how to use CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).
