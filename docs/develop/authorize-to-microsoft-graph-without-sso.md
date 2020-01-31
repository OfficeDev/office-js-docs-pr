---
title: Authorize to Microsoft Graph without SSO
description: 'Learn to authorize to Microsoft Graph without SSO'
ms.date: 01/29/2020
localization_priority: Priority
---

# Authorize to Microsoft Graph without SSO

Your add-in can get authorization to Microsoft Graph data by obtaining an access token to Graph from Azure Active Directory (AAD). Use either the Authorization Code flow or the Implicit flow just as you would in other web applications but with one exception: AAD does not allow its login page to open in an iframe. When an Office Add-in is running on *Office on the web*, the task pane is an iframe. This means you'll need to open the AAD login screen in a dialog box opened with the Office dialog API. This affects how you use authentication and authorization helper libraries. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

For information about programming authentication with AAD, begin with [Microsoft identity platform (v2.0) overview](/azure/active-directory/develop/v2-overview), where you'll find tutorials and guides in that documentation set, as well as links to relevant samples. Once again, you may need to adjust the code in the samples to run in the Office dialog box to account for the fact that the Office dialog box runs in a separate process from the task pane.

After your code obtains the access token to Graph, either it passes the access token from the dialog box to the task pane, or it stores the token in a database and signals the task pane that the token is available. (See [Authentication with the Office dialog API](auth-with-office-dialog-api.md) for details.) Code in the task pane requests data from Graph, and includes the token in those requests. For more information about calling Graph and the Graph SDKs, see [Microsoft Graph documentation](/graph/).

## Recommended libraries and samples

We recommend that you use the following libraries when accessing Microsoft Graph without using SSO:

- For add-ins using a server-side with a .NET-based framework such as .NET Core or ASP.NET, use [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- For add-ins using a NodeJS-based server-side, use [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- For add-ins using the Implicit flow, use [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

For more information about recommended libraries for working with Microsoft Identity Platform (formerly AAD v.2.0), see [Microsoft identity platform authentication libraries](/azure/active-directory/develop/reference-v2-libraries).

The following samples get Microsoft Graph data from an Office Add-in:

- [Office Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
