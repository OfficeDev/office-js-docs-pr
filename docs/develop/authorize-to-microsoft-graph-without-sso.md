---
title: Authorize to Microsoft Graph without SSO
description: ''
ms.date: 08/07/2019
localization_priority: Priority
---

# Authorize to Microsoft Graph without SSO

You can get authorization to Microsoft Graph data for your add-in by obtaining an access token to Graph from Azure Active Directory (AAD). You do this using either the Authorization Code flow or the Implicit Flow just as you would any in any other web application with one exception: AAD does not allow its login page to open in an iframe. When an Office Add-in is running on *Office on the web*, the task pane is an iframe. This means that you will need to open the AAD login screen in a dialog opened with the Office Dialog API. This will effect how you use authentication and authorization helper libraries. For more information, see [Authentication with the Office Dialog API](auth-with-office-dialog-api.md).

For information about programming authentication with AAD, begin with [Microsoft identity platform (v2.0) overview](/azure/active-directory/develop/v2-overview). There are many tutorials and guides in that documentation set, as well as links to relevant samples. Once again a reminder: you may need to adjust the code in the samples to run in the Office Dialog to account for the fact that the Dialog runs in a separate process from the task pane.

After your code has obtained the access token to Microsoft Graph, either it passes the access token from the dialog to the task pane, or it stores the token in a database and signals the task pane that the token is available there. (See [Authentication with the Office Dialog API](auth-with-office-dialog-api.md) for details.) Code in the task pane requests data from Microsoft Graph, and includes the token in those requests. For more information about calling Microsoft Graph and the SDKs for Microsoft Graph, see [Microsoft Graph documentation](/graph/).

## Recommended libraries and samples

We recommend that you use the following libraries when accessing Microsoft Graph without using SSO:

- For add-ins using a server-side with a .NET-based framework such as .NET Core or ASP.NET, use [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- For add-ins using a NodeJS-based server-side, use [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- For add-ins using the Implicit flow, use [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

For more information about recommended libraries for working with Microsoft Identity Platform (formerly AAD v.2.0), see [Microsoft identity platform authentication libraries](/azure/active-directory/develop/reference-v2-libraries).

The following samples get Microsoft Graph data from an Office Add-in:

- [Office Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/office-add-in-microsoft-graph-aspnet)
- [Outlook Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/outlook-add-in-microsoft-graph-aspnet)

