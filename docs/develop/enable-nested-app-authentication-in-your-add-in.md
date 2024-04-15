---
title: Enable SSO in an Office Add-in using nested app authentication
description: Learn how to enable SSO in an Office Add-in using nested app authentication.
ms.date: 04/12/2024
ms.topic: how-to
ms.localizationpriority: high
---

# Enable SSO in an Office Add-in using nested app authentication (preview)

You can use the MSAL.js library (version 3.11 and later) with nested app authentication to use SSO from your Office Add-in. Using nested app authentication offers several advantages over the On-Behalf-Of (OBO) flow.

- You only need to use the MSAL.js library and don’t need the `getAccessToken` function in Office.js.
- You can call services such as Microsoft Graph with an access token from your client code as an SPA. There’s no need for a middle-tier server.
- You can use incremental and dynamic consent for scopes.
- You don't need to [preauthorize your hosts](/microsoftteams/platform/m365-apps/extend-m365-teams-personal-tab?tabs=manifest-teams-toolkit#update-azure-ad-app-registration-for-sso) (For example, Teams, Office) to call your endpoints.

> [!IMPORTANT]
> Nested app authentication (NAA) is currently in preview. To try this feature, join the Microsoft 365 Insider Program (https://insider.microsoft365.com/join) and choose the Beta Channel. Don't use NAA in production add-ins. We invite you to try out NAA in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).
> NAA is supported in the following builds.
>
> - Word, Excel, and PowerPoint on Windows build 16.0.17531.20000 or later.
> - Word, Excel, and PowerPoint on Mac build 16.85.24040319 or later.
> - Outlook on Windows build 16.0.17531.20000 or later.
> - Outlook on Mac build 16.85.24040319 or later.

## Register your single-page application

You’ll need to create a Microsoft Azure App registration for your add-in on the Azure portal. The app registration must have at minimum:

- A name
- A supported account type
- An SPA redirect

If your add-in requires additional app registration beyond NAA and SSO, see [Single-page application: App registration](/entra/identity-platform/scenario-spa-app-registration).

## Add a trusted broker through SPA redirect

To enable NAA, your app registration must include a specific redirect URI to indicate to the Microsoft identity platform that your add-in allows itself to be brokered by supported hosts. The redirect URI of the application must be of type **Single Page Application** and conform to the following scheme.

`brk-multihub://your-add-in-domain`

Trusted broker groups are dynamic by design and can be updated in the future to include additional hosts where your add-in may use NAA flows. Currently the brk-multihub group includes Office Word, Excel, PowerPoint, Outlook, and Teams (for when Office is activated inside).

## Configure MSAL config to use NAA

Configure your add-in to use NAA by setting the `supportsNestedAppAuth` property to true in your MSAL configuration. This enables MSAL to use APIs on its native application host (for example, Outlook) to acquire tokens for your application. If you don't set this property, MSAL uses the default JavaScript-based implementation to acquire tokens for your application, which may lead to unexpected auth prompts and unsatisfiable conditional access policies when running inside of a webview.

The following steps show how to enable NAA in the `taskpane.js` or `taskpane.ts` file in a project built with `yo office`. The code can be adapted to any project.

1. Add the `@azure/msal-browser` package to the `dependencies` section of the `package.json` file for your project.

    ```json
    "dependencies": {
        "@azure/msal-browser": "^3.11.1",
        ...
    ```

1. Save and run `npm install` to install `@azure/msal-browser`.

1. Add the following code to the top of the `taskpane.js` or `taskpane.ts` file. Replace the `Enter_the_Application_Id_Here` placeholder with the Azure app ID you saved previously.

    ```JavaScript
    import { PublicClientNext } from "@azure/msal-browser";
    
    // Configuration for NAA.  
    const msalConfig = { 
      auth: { 
        clientId: "Enter_the_Application_Id_Here", 
        authority: "https://login.microsoftonline.com/common", 
        supportsNestedAppAuth: true 
      } 
    }
    let pca = undefined; // public client application to be initialized later.
    ```

## Initialize the public client application

Next, you need to initialize MSAL and get an instance of the public client application. This is used to get access tokens when needed. It's recommended to create the public client application in the `Office.onReady` method.

- In your `Office.onReady` function, add a call to `createPublicClientApplication` as shown below to initilize the `pca` variable.

    ```javascript
    Office.onReady(async (info) => {
      if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
        
        // Initialize the publice client application
        pca = await PublicClientNext.createPublicClientApplication(msalConfig);
      }
    });
    ```

## Acquire your first token

The tokens acquired by MSAL.js via NAA will be issued for your Azure app registration ID. In this code sample, you acquire a token for the Microsoft Graph API. If the user has an active session with Microsoft Entra ID the token is acquired silently. If not, the library prompts the user to sign in interactively. The token is then used to call the Microsoft Graph API.

The following steps show the pattern to use for acquiring a token.

1. Specify your scopes. NAA supports incremental and dynamic consent so always request the minimum scopes needed for your code to complete its task.
1. Call `acquireTokenSilent`. This will get the token without requiring user interaction.
1. If `acquireTokenSilent` fails, call `acquireTokenPopup` to display an interactive dialog for the user. `acquireTokenSilent` can fail if the token expired, or the user has not yet consented to all of the requested scopes.

The following code shows how to implement this authentication pattern in your own project.

1. Add the following code to `taskpane.js` or `taskpane.ts`. The `getFileNames` function specifies the scopes needed when it calls the `ssoGetToken` function.

    ```javascript
    /**
     * Gets the first 10 (or specified number) file names from
     * the user's OneDrive account.
     * 
     * @param {*} count Number of file names to get.
     * @returns 
     */
    async function getFileNames(count=10) {
      const accessToken = await ssoGetToken(["Files.Read","User.Read","openid","profile"]);
      const response = await makeGraphRequest(
        accessToken,
        "/me/drive/root/children",
        `?$select=name&$top=${count}`
      );
    
      const names = response.value.map((item) => item.name);
      return names;
    }
    ```

1. Next, add the following `ssoGetToken` function. This function calls `acquireTokenSilent` to get the access token. If that fails, it will call `acquireTokenPopup` to get the access token interactively.

    ```JavaScript
    /**
     * Uses MSAL and nested app authentication to get an access token using SSO.
     * 
     * @param {*} scopes Minimal scopes required for the token.
     * @returns An access token for user signed in to Office.
     */
    async function ssoGetToken(scopes) {
      if (pca === undefined) {
        throw new Error("AccountManager is not initialized!");
      }
    
      // Specify minimum scopes needed for the access token.
      const tokenRequest = {
        scopes: scopes
      };
    
      try {
        console.log("Trying to acquire token silently...");
        const userAccount = await pca.acquireTokenSilent(tokenRequest);
        console.log("Acquired token silently.");
        return userAccount.accessToken;
      } catch (error) {
        console.log(`Unable to acquire token silently: ${error}`);
      }
    
      // Acquire token silent failure. Send an interactive request via popup.
      try {
        console.log("Trying to acquire token interactively...");
        const userAccount = await pca.acquireTokenPopup(tokenRequest);
        console.log("Acquired token interactively.");
        return userAccount.accessToken;
      } catch (popupError) {
        // Acquire token interactive failure.
        console.log(`Unable to acquire token interactively: ${popupError}`);
        throw new Error(`Unable to acquire access token: ${popupError}`);
      }
    }
    ```

## Call an API

After acquiring the token, use it to call an API. The following example shows how to call the Microsoft Graph API by calling `fetch` with the token attached in the *Authorization* header.

1. Add the following code to `taskpane.js` or `taskpane.ts`. The `makeGraphRequest` function calls the specified REST API and returns the result.

    ```javascript
    /**
     *  Calls a Microsoft Graph API and returns the response.
     *
     * @param accessToken The access token to use for the request.
     * @param path Path component of the URI, e.g., "/me". Should start with "/".
     * @param queryParams Query parameters, e.g., "?$select=name,id". Should start with "?".
     * @returns
     */
    async function makeGraphRequest(accessToken, path, queryParams) {
      if (!path) throw new Error("path is required.");
      if (!path.startsWith("/")) throw new Error("path must start with '/'.");
      if (queryParams && !queryParams.startsWith("?")) throw new Error("queryParams must start with '?'.");
    
      const response = await fetch(`https://graph.microsoft.com/v1.0${path}${queryParams}`, {
        headers: { Authorization: accessToken },
      });
    
      if (response.ok) {
        const data = await response.json();
        return data;
      } else {
        throw new Error(response.statusText);
      }
    }
    ```

1. Use the following code to replace the `run` function to call `getFileNames`. This code writes the file names to the debug console. You need to ensure this function is called by a button on the task pane.

    ```javascript
    async function run() {
      try {
        const fileNames = await getFileNames();
        fileNames.forEach(fileName => {
          console.log(fileName);
        })
      } catch (error) {
        console.error(error);
      }
    }
    ```

## What is nested app authentication

Nested app authentication enables SSO for applications that are nested inside of supported Microsoft applications. For example, Excel on Windows runs your add-in inside a webview. In this scenario, your add-in is a nested application running inside Excel, which is the host. NAA also supports nested apps in Teams. For example, if a Teams tab is hosting Excel, and your add-in is loaded, it is nested inside Excel, which is also nested inside Teams. Again, NAA supports this nested scenario and you can access SSO to get user identity and access tokens of the signed in user.

## NAA supported accounts and hosts

NAA supports both Microsoft Accounts and Microsoft Entra ID (work/school) identities. It doesn’t support B2C scenarios. For preview, NAA is supported in Office on Windows and Mac. For GA, NAA will also support Office on the web, iOS, and Outlook Mobile on Android and iOS.

## Best practices

We recommend the following best practices when using MSAL.js with NAA.

### Use silent authentication whenever possible

MSAL.js provides the `acquireTokenSilent` method that handles token renewal by making silent token requests without prompting the user. The method first looks for a valid cached token. If it doesn't find one, the library makes the silent request to Microsoft Entra ID and if there's an active user session, a fresh token is returned.  

In certain cases, the `acquireTokenSilent` method's attempt to get the token fails. Some examples of this are when there's an expired user session with Microsoft Entra ID or a password change by the user, which requires user interaction. When the acquireTokenSilent fails, you need to call the interactive `acquireTokenPopup` token method.

### Have a fallback when NAA isn't supported

While we strive to provide a high-degree of compatibility with these flows across the Microsoft ecosystem, your add-in may be loaded in an older Office host that does not support NAA. In these cases, your add-in won't support seamless SSO and you may need to fall back to an alternate method of authenticating the user. In generaly you'll want to use the MSAL SPA authentication pattern with the Office JS dialog API. For more information, see the following resources.

- [Authenticate and authorize with the Office dialog API](/office/dev/add-ins/develop/auth-with-office-dialog-api).
- [Microsoft identity sample for SPA and JavaScript](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md)
- [Microsoft identity samples for various app types and frameworks](https://learn.microsoft.com/entra/identity-platform/sample-v2-code?tabs=apptype)

## MSAL.js APIs supported by NAA

The following table shows which APIs are supported when NAA is enabled in the MSAL config.

| Method                        | Supported by NAA |
|-------------------------------|------------------|
| *acquireTokenByCode*          | NO (throws exception) |
| *acquireTokenPopup*           | YES              |
| *acquireTokenRedirect*        | NO (throws exception) |
| *acquireTokenSilent*          | YES              |
| *addEventCallback*            | YES              |
| *addPerformanceCallback*      | NO (throws exception) |
| *disableAccountStorageEvents* | NO (throws exception) |
| *enableAccountStorageEvents*  | NO (throws exception) |
| *getAccountByHomeId*          | YES              |
| *getAccountByLocalId*         | YES              |
| *getAccountByUsername*        | YES              |
| *getActiveAccount*            | YES              |
| *getAllAccounts*              | YES              |
| *getConfiguration*            | YES              |
| *getLogger*                   | YES              |
| *getTokenCache*               | NO (throws exception) |
| *handleRedirectPromise*       | NO               |
| *initialize*                  | YES              |
| *initializeWrapperLibrary*    | YES              |
| *loginPopup*                  | YES              |
| *loginRedirect*               | NO (throws exception) |
| *logout*                      | NO (throws exception) |
| *logoutPopup*                 | NO (throws exception) |
| *logoutRedirect*              | NO (throws exception) |
| *removeEventCallback*         | YES              |
| *removePerformanceCallback*   | NO (throws exception) |
| *setActiveAccount*            | NO               |
| *setLogger*                   | YES              |
| *ssoSilent*                   | YES              |

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft.com](mailto:secure@microsoft.com) with as much detail as you can provide. Your submission may be eligible for a bounty through the [Microsoft Bounty](https://aka.ms/bugbounty) program. Don't post security issues to GitHub or any other public site. We'll contact you shortly after receiving your issue report. We encourage you to get new security incident notifications by visiting [Microsoft technical security notifications](https://technet.microsoft.com/security/dd252948) to subscribe to Security Advisory Alerts.

## Code samples

|Sample name | Description |
|----------------|--------------------------------------------------------|
| [Office Add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-SSO-NAA) | Shows how to use MSAL.js nested app authentication (NAA) in an Office Add-in to access Microsoft Graph APIs for the signed in user. The sample displays the signed in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into the document. |
| [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA) | Shows how to use MSAL.js nested app authentication (NAA) in an Outlook Add-in to access Microsoft Graph APIs for the signed in user. The sample displays the signed in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into a new message body.|
