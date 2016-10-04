# Authorize external services in your Office Add-in

Popular online services, including Office 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, enable developers to give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in. 

The industry standard framework for enabling web application access to an online service is called OAuth 2.0. In most situations, you don't need to know the details of how the framework works to make use of it in your add-in. Many libraries are available that abstract the details for you.

A fundamental idea of OAuth is that an application can be a security principal unto itself, just like a user or a group, with its own identity and set of permissions. In the most typical scenarios, when the user takes an action in the Office add-in that requires the online service, the add-in sends the service a request for a specific set of permissions to the user's account. The service then prompts the user to grant the add-in those permissions. After the permissions are granted, the service sends the add-in a small encoded *access token*. The add-in can use the service by including the token in all its requests to the service's APIs. But the add-in can act only within the permissions that the user granted it. The token also expires after a specified time.

Several OAuth patterns, called *flows* or *grant types*, are designed for different scenarios. The following are the two most important:

- **Implicit flow**: Communication between the add-in and the online service is implemented with client-side JavaScript.
- **Authorization Code flow**: Communication is *server-to-server* between your add-in's web application and the online service. So, it is implemented with server-side code.

The purpose of the flows is to secure the identity and authorization of the application. In the Authorization Code flow, you are provided a *client secret* that needs to be kept hidden. A Single Page Application (SPA) has no way to protect the secret, so we recommend that you use the Implicit flow in SPAs. 

You should be familiar with the other pros and cons of the two flows. The official definitions at [Authorization Code](https://tools.ietf.org/html/rfc6749#section-1.3.1) and [Implicit](https://tools.ietf.org/html/rfc6749#section-1.3.2) are a good starting place. 

>**Note:** You also have the option of having a middleman service do all the authorizing for you and passing the access token to your add-in. For details, see the section *Middleman services* later in this article.

## Using the Implicit flow in Office Add-ins
The best way to find out if the online service supports the Implicit flow is to consult the documentation.

For services that support it, we provide a JavaScript library that does all the detailed work for you:

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

The \demo folder of the repo contains a sample add-in that uses the library to access some popular services including Google, Facebook, and Office 365.

See also the **Libraries** section later in this article.

## Using the Authorization Code flow in Office Add-ins

We have some sample add-ins that use the Authorization Code flow:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For details, see the **Libraries** section later in this article.

### Relay/Proxy functions

You can use the Authorization Code flow even with a serverless web application by storing the *client ID* and *client secret* values in a simple function that is hosted in a service such as [Azure Functions](https://azure.microsoft.com/en-us/services/functions) or [Amazon Lambda](https://aws.amazon.com/lambda).
The function exchanges a given code for an appropriate *access token* and relays it back to the client. The security of this approach depends on how well access to the function is guarded.

To use this technique, your add-in displays a UI/popup to show the login screen for the online service (Google, Facebook, and so on). When the user is logged on and grants the add-in permission to her resources in the online service, the developer receives a code which can be then sent to the online function. The services described in **Middleman services** in this article use a flow similar to this. 

## Libraries

Libraries are available for many languages and platforms, and for both flows. Some are general purpose, others are for specific online services. 

**Office 365 and other services that use Azure Active Directory as the authorization provider**: [Azure Active Directory Authentication Libraries](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). A preview is also available for the [Microsoft Authentication Library](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

**Facebook**: Search [Facebook for Developers](https://developers.facebook.com) for "library" or "sdk". 

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](http://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## Middleman services

Your add-in can use a middleman service, such as Auth0, that either provides access tokens for many popular online services, or simplifies the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman and it will send back any required tokens for the online service. All the authorization implementation code is in the middleman service. 

We have a sample that uses Auth0 to enable social login with Facebook, Google, and Microsoft Accounts:

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## What is CORS?

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS). For information about how you can use work with CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations).
