---
title: Authorize to Microsoft Graph from an Office Add-in
description: Learn to authorize to Microsoft Graph from an Office Add-in.
ms.date: 01/12/2026
ms.localizationpriority: medium
---

# Authorize to Microsoft Graph from an Office Add-in

Your add-in can get authorization to Microsoft Graph data by obtaining an access token to Microsoft Graph from the Microsoft identity platform. Use either the Authorization Code flow or the Implicit flow just as you would in other web applications but with one exception: The Microsoft identity platform doesn't allow its sign-in page to open in an iframe. When an Office Add-in is running in *Office on the web*, the task pane is an iframe. This means you'll need to open the sign-in page in a dialog box by using the Office dialog API. This affects how you use authentication and authorization helper libraries. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

> [!NOTE]
> If you're implementing SSO and plan to access Microsoft Graph, see [Enable single sign-on in an Office Add-in with nested app authentication](enable-nested-app-authentication-in-your-add-in.md).

For information about programming authentication using the Microsoft identity platform, see [Microsoft identity platform documentation](/azure/active-directory/develop). You'll find tutorials and guides in that documentation set, as well as links to relevant samples. Once again, you may need to adjust the code in the samples to run in the Office dialog box to account for the Office dialog box that runs in a separate process from the task pane.

After your code obtains the access token to Microsoft Graph, it passes the access token from the dialog box to the task pane. For details, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md). Code in the task pane requests data from Microsoft Graph and includes the token in those requests. For more information about calling Microsoft Graph and the Microsoft Graph SDKs, see [Microsoft Graph documentation](/graph/).

## Recommended libraries and samples

The following libraries provide access to Microsoft Graph.

- For add-ins using a server-side with a .NET-based framework such as .NET Core or ASP.NET, use [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- For add-ins using a NodeJS-based server-side, use [Microsoft Authentication Library for Node](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/README.md).

For more information about recommended libraries for working with Microsoft Identity Platform, see [Microsoft identity platform authentication libraries](/azure/active-directory/develop/reference-v2-libraries).

The following samples get Microsoft Graph data from an Office Add-in.

- [Office Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)

[!INCLUDE [chrome-tracking-prevention](../includes/chrome-tracking-prevention.md)]
