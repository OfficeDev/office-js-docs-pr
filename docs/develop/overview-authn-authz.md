---
title: Overview of authentication and authorization in Office Add-ins
description: 'Require users to authenticate login in Web applications and Office Add-ins.'
ms.date: 07/07/2020
localization_priority: Priority
---

# Overview of authentication and authorization in Office Add-ins

Web applications and, hence, Office Add-ins allow anonymous access by default, but you can require users to authenticate with a login. For example, you can require that your users be logged in with a Microsoft account, a Microsoft 365 Education or work account, or other common account. This task is called user authentication because it enables the add-in to know who the user is.

Your add-in can also get the user's consent to access their Microsoft Graph data (such as their Microsoft 365 profile, OneDrive files, and SharePoint data) or to data in other external sources such as Google, Facebook, LinkedIn, SalesForce, and GitHub. This task is called add-in (or app) authorization, because it is the *add-in* that is being authorized, not the user.

You have a choice of two ways to accomplish these authentications.

- **Office Single Sign-on (SSO)**: A system, *currently in preview*, that enables the user's login to Office to also function as a login to the add-in. Optionally, the add-in can also use the user's Office credentials to authorize the add-in to Microsoft Graph. (Non-Microsoft sources are not accessible through this system.)
- **Web Application Authentication and Authorization with Azure Active Directory**: This isn't something new or special. It's just the way Office add-in (and other web apps) authenticated users and authorized apps before there was an Office SSO system and is still used in scenarios where Office SSO cannot be.

The following flowchart shows you the decisions that you need to make as an add-in developer. Details are later in this article.

![An image showing a decision flowchart for enabling authentication and authorization in Office Add-ins](../images/authflowchart.png)

## User authentication without SSO

You can authenticate a user in an Office Add-in with Azure Active Directory (AAD) as you would in other web applications with one exception: AAD does not allow its login page to open in an iframe. When an Office Add-in is running on *Office on the web*, the task pane is an iframe. This means that you'll need to open the AAD login screen in a dialog box opened with the Office dialog API. This affects how you use authentication helper libraries. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

For information about programming authentication with AAD, begin with [Microsoft identity platform (v2.0) overview](/azure/active-directory/develop/v2-overview) where you'll find many tutorials and guides, as well as links to relevant samples and libraries. As explained in [Authentication with the Office dialog API](auth-with-office-dialog-api.md), you may need to adjust the code in the samples to run in the Office dialog box.

## Access to Microsoft Graph without SSO

You can get authorization to Microsoft Graph data for your add-in by obtaining an access token to Graph from Azure Active Directory (AAD). You can do this without relying on Office SSO. For more information about how, see [Access to Microsoft Graph without SSO](authorize-to-microsoft-graph-without-sso.md) which has more details and links to samples.

## User authentication with SSO

To authenticate the user using SSO, your code in a task pane or function file calls the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) method. If the user is not signed in, Office will open a dialog box and navigate to the Azure Active Directory login page. After the user signs in, or if the user is already signed in, the method returns an access token. The token is a bootstrap token in the **On Behalf Of** flow. (See [Access to Microsoft Graph with SSO](#access-to-microsoft-graph-with-sso).) However, it can be used as an ID token as well, because it contains several claims that are unique to the current user, including `preferred_username`, `name`, `sub`, and `oid`. For guidance on which property to use as the ultimate user ID, see [Microsoft identity platform access tokens](https://docs.microsoft.com/azure/active-directory/develop/access-tokens#payload-claims). For an example of a one of these tokens, see [Example access token](sso-in-office-add-ins.md#example-access-token).

After your code has extracted the desired claim from the token, it uses that value to look up the user in a user table or user database that you maintain. Use the database to store user-relative information such as the user's preferences or the state of the user's account. Since you are using SSO, your users don't sign-in separately to your add-in, so you do not need to store a password for the user.

Before you begin implementing user authentication with SSO, be sure that you are thoroughly familiar with the article [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md). Note also these samples:

- [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especially the file [ssoAuthES6.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js).
- [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).

These samples, however, do not use the token as an ID token. They use it to get access to Microsoft Graph with the **On Behalf Of** flow.

## Access to Microsoft Graph with SSO

To use SSO to access Microsoft Graph, your add-in in a task pane or function file calls the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) method. If the user is not signed in, Office will open a dialog box and navigate it to the Azure Active Directory login page. After the user signs in, or if the user is already signed in, the method returns an access token. The token is a bootstrap token in the **On Behalf Of** flow. Specifically, it has a `scope` claim with the value `access_as_user`. For guidance about the claims in the token, see [Microsoft identity platform access tokens](https://docs.microsoft.com/azure/active-directory/develop/access-tokens#payload-claims). For an example of a one of these tokens, see [Example access token](sso-in-office-add-ins.md#example-access-token).

After your code obtains the token, it uses it in the **On Behalf Of** flow to obtain a second token: an access token to Microsoft Graph.

Before you begin implementing Office SSO, be sure that you are thoroughly familiar with these two articles:

- [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md)
- [Authorize to Microsoft Graph with SSO](authorize-to-microsoft-graph.md)

You should also read at least one of the walkthrough articles listed here. Even if you don't carry out the steps, these contain valuable information about how you implement Office SSO and the **On Behalf Of** flow. 

- [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md)
- [Create an Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md)

Note also these samples:

- [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)

## Access to non-Microsoft data sources

Popular online services, including Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in. For an overview of the ways that your add-in can do this, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).

> [!IMPORTANT]
> Before you begin coding, find out if the data source allows its login in screen to open in an iframe. When an Office Add-in is running on *Office on the web*, the task pane is an iframe. If the data source does not allow its login screen to open in an iframe, then you'll need to open the login screen in a dialog box opened with the Office dialog API. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).
