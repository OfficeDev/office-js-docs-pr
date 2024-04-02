---
title: Enable SSO in an Office Add-in using nested app authentication
description: Learn how to enable SSO in an Office Add-in using nested app authentication.
ms.date: 03/20/2024
ms.topic: how-to
ms.localizationpriority: medium

---

# Enable SSO in an Office Add-in using nested app authentication (preview)

You can use the MSAL.js library (version 3.10 and later) with nested app authentication to use SSO from your Office Add-in. Using nested app authentication offers several advantages over the On-Behalf-Of (OBO) flow.

- You only need to use the MSAL.js library and don’t need to use the `getAccessToken` function in Office.js.
- You can call services such as Microsoft Graph with an access token from your client code as an SPA. There’s no need for a middle-tier server.
- You can use incremental and dynamic consent for scopes.
- You don't need to [preauthorize your hosts](/microsoftteams/platform/m365-apps/extend-m365-teams-personal-tab?tabs=manifest-teams-toolkit#update-azure-ad-app-registration-for-sso) (For example, Teams, Office) to call your endpoints.

> [!IMPORTANT]
> Nested app authentication is currently in preview. To try this feature you need to join the Microsoft 365 Insider Program (https://insider.microsoft365.com/en-us/join) and choose the Beta Channel. Don't use NAA in production add-ins. We invite you to try out NAA in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

## Register your single-page application

You’ll need to create a Microsoft Azure App registration for your add-in on the Azure portal. The app registration must have at minimum:

- A name.
- A supported account type.
- An SPA redirect.

If your add-in requires additional app registration beyond NAA and SSO, see [Single-page application: App registration](/entra/identity-platform/scenario-spa-app-registration)

## Add a trusted broker through SPA redirect

To enable NAA, your app registration must include a specific redirect URI to indicate to the Microsoft identity platform that your add-in allows itself to be brokered by supported hosts. The redirect URI of the application must be of type **Single Page Application** and conform to the following scheme.

`brk-multihub://your-add-in-domain`

Trusted broker groups are dynamic by design and can be updated in the future to include additional hosts where your add-in may use NAA flows. Currently the brk-multihub group includes Office Word, Excel, PowerPoint, Outlook, and Teams (for when Office is activated inside).

## Configure MSAL config to use NAA

Configure your add-in to use NAA by setting the `supportsNestedAppAuth` property to true in your MSAL configuration. This enables MSAL to use APIs on its native application host (For example, Outlook) to acquire tokens for your application. If you don't set this property, MSAL uses the default JavaScript-based implementation to acquire tokens for your application, which may lead to unexpected auth prompts and unsatisfiable conditional access policies when running inside of a webview.

```JavaScript
// Configuration for NAA.  

const msalConfig = { 
  auth: { 
    clientId: "Enter_the_Application_Id_Here", 
    authority: "https://login.microsoftonline.com/common", 
    supportsNestedAppAuth: true 
  } 
} 
```

## Initialize the public client application

Next you need to initialize MSAL and get an instance of the public client application. This is used to get access tokens when needed. It's recommended to create the public client application in the `Office.onReady` method.

```javascript
let pca = undefined;

// Initialize the publice client application
Office.onReady(async (info) => {
    pca = await msalBrowser.PublicClientNext.createPublicClientApplication(msalConfig);
  });
```

## Acquire your first token

The tokens acquired by MSAL.js via NAA will be issued for your Azure app registration ID. In this code sample, we acquire a token for the Microsoft Graph API. The token is acquired silently if the user has an active session with Microsoft Entra ID. If not, the library prompts the user to sign in interactively. The token is then used to call the Microsoft Graph API.

The following steps show the pattern to use for acquiring a token:

1. Specify your scopes. NAA supports incremental and dynamic consent so always request the minimum scopes needed for your code to complete its task.
1. Call `acquireTokenSilent`. This will get the token without requiring user interaction.
1. If `acquireTokenSilent` fails, call `acquireTokenPopup` to display an interactive dialog for the user. `acquireTokenSilent` can fail if the token expired, or the user has not yet consented to all of the requested scopes.

```JavaScript
async function run() { 
  // Specify minimum scopes needed for the access token. 
  const tokenRequest = { 
    scopes: ["User.Read", "openid", "profile"], 
    loginHint: myloginHint 
  } 
 
  try { 
    const userAccount = await pca.acquireTokenSilent(tokenRequest); 
    // Call your API with the token. 
    makeMSGraphCall(userAccount.accessToken); 
  } catch (error) { 
    // Acquire token silent failure. Send an interactive request via popup. 
    try { 
      const userAccount = pca.acquireTokenPopup(tokenRequest); 
      // Call  your API with the token. 
      makeMSGraphCall(userAccount.accessToken); 
    } catch (popupError) { 
      // Acquire token interactive failure. 
      console.log(popupError); 
    } 
  } 
} 
```

## Call an API

After acquiring the token use it to call an API. The following example shows how to call the Microsoft Graph API by calling `fetch` with the token attached in the *Authorization* header.

```javascript
async function makeMSGraphCall(accessToken) {
  const requestString = "https://graph.microsoft.com/v1.0/me";
  const headersInit = { 'Authorization': accessToken };
  const requestInit = { 'headers': headersInit }

  // Make REST call to MS Graph.
  const result = await fetch(requestString, requestInit);
  if (result.ok) {
    const data = await result.text();
    console.log(data);
    document.getElementById("userInfo").innerText = data;
  } else {
    // Likely an MS Graph error if result was not ok.
    // Error details are in the body.
    const response = await result.text();
    console.log(response);
  }
}
```

## What is nested app authentication

Nested app authentication enables SSO for applications that are nested inside of supported first-party applications. For example, Excel on Windows runs your add-in inside a webview. In this scenario, your add-in is a nested application running inside Excel, which is the host. NAA also supports nested apps in Teams. For example, if a Teams tab is hosting Excel, and your add-in is loaded, it is nested inside Excel, which is also nested inside Teams. Again, NAA supports this nested scenario and you can access SSO to get user identity and access tokens of the signed in user.

## NAA supported accounts and hosts

NAA supports both Microsoft Accounts and Microsoft Entra ID (work/school) identities. It doesn’t support B2C scenarios. For preview, NAA is supported in Office on Windows and Mac. For GA, NAA will also support Office on the web, iOS, and Outlook Mobile on Android and iOS.

## Best practices

The following are some best practices to follow when using MSAL.js with NAA.

### Use silent authentication whenever possible

MSAL.js provides the `acquireTokenSilent` method that handles token renewal by making silent token requests without prompting the user. The method first looks for a valid cached token. If it doesn't find one, the library makes the silent request to Microsoft Entra ID and if there's an active user session, a fresh token is returned.  

In certain cases, the `acquireTokenSilent` method's attempt to get the token fails. Some examples of this are when there's an expired user session with Microsoft Entra ID or a password change by the user, and so on. which requires user interaction. When the acquireTokenSilent fails, you need to call the interactive acquire token method (acquireTokenPopup).

### Have a fallback when NAA isn't supported

While we strive to provide a high-degree of compatibility with these flows across the Microsoft ecosystem, your application may appear in downlevel/legacy clients that haven't been updated to support NAA. In these cases, your application won't support seamless SSO and you may need to invoke special APIs for interacting with the user to open authentication dialogs. For more information, see [Authenticate and authorize with the Office dialog API](/office/dev/add-ins/develop/auth-with-office-dialog-api).

## MSAL.js APIs supported by NAA

The following table shows which methods are supported when NAA is enabled in the MSAL config.

| Method                        | Supported by NAA |
|-------------------------------|------------------|
| *acquireTokenByCode*          | NO (throws exception)               |
| *acquireTokenPopup*           | YES              |
| *acquireTokenRedirect*        | NO (throws exception)              |
| *acquireTokenSilent*          | YES              |
| *addEventCallback*            | YES              |
| *addPerformanceCallback*      | YES              |
| *disableAccountStorageEvents* | NO (throws exception)               |
| *enableAccountStorageEvents*  | NO (throws exception)             |
| *getAccountByHomeId*          | YES              |
| *getAccountByLocalId*         | YES              |
| *getAccountByUsername*        | YES              |
| *getActiveAccount*            | YES              |
| *getAllAccounts*              | YES              |
| *getConfiguration*            | YES              |
| *getLogger*                   | YES              |
| *getTokenCache*               | NO (throws exception)              |
| *handleRedirectPromise*       | NO               |
| *initialize*                  | YES              |
| *initializeWrapperLibrary*    | YES              |
| *loginPopup*                  | YES              |
| *loginRedirect*               | NO (throws exception)              |
| *logout*                      | NO (throws exception)              |
| *logoutPopup*                 | NO (throws exception)              |
| *logoutRedirect*              | NO (throws exception)              |
| *removeEventCallback*         | YES              |
| *removePerformanceCallback*   | YES               |
| *setActiveAccount*            | NO               |
| *setLogger*                   | YES               |
| *ssoSilent*                   | YES              |

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft)](secure@microsoft).com with as much detail as you can provide. Your submission may be eligible for a bounty through the [Microsoft Bounty](https://aka.ms/bugbounty) program. Don't post security issues to GitHub or any other public site. We'll contact you shortly after receiving your issue report. We encourage you to get new security incident notifications by visiting [Microsoft technical security notifications](https://technet.microsoft.com/security/dd252948) to subscribe to Security Advisory Alerts.
