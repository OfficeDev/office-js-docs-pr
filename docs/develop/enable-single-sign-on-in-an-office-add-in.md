---
title: Enable single sign-on in an Office Add-in with nested app authentication
description: Learn how to enable SSO in an Office Add-in with nested app authentication.
ms.date: 11/24/2025
ms.topic: how-to
ms.localizationpriority: high
---

# Enable single sign-on in an Office Add-in with nested app authentication

Use the MSAL.js library with nested app authentication (NAA) to enable single sign-on (SSO) from your Office Add-in. The procedures in this article guide you through creating an app registration and adding code to your project to use NAA successfully.

## Register your single-page application

You’ll need to create a Microsoft Azure App registration for your add-in on the Azure portal. The app registration must have at minimum:

- A name
- A supported account type
- An SPA redirect for NAA and for your task pane

If your add-in requires additional app registration beyond NAA and SSO, see [Single-page application: App registration](/entra/identity-platform/scenario-spa-app-registration).

## Add a trusted broker through SPA redirect

To enable NAA, your app registration must include a specific redirect URI to indicate to the Microsoft identity platform that your add-in allows itself to be brokered by supported hosts. The redirect URI of the application must be of type **Single Page Application** and conform to the following scheme.

`brk-multihub://your-add-in-domain`

Your domain must include only the origin and not its subpaths. For example:

✔️ brk-multihub://localhost:3000<br>
✔️ brk-multihub://www.contoso.com<br>
❌ brk-multihub://www.contoso.com/go

Trusted broker groups are dynamic by design and can be updated in the future to include additional hosts where your add-in may use NAA flows. Currently the brk-multihub group includes Office Word, Excel, PowerPoint, Outlook, and Teams (for when Office is activated inside).

> [!IMPORTANT]
> For Word, Excel, and PowerPoint in the browser, you also need an additional redirect since the browser uses a standard authentication flow. The SPA redirect URI must reference the HTML page where you will use the MSAL.js library to request tokens through NAA.

Use the following steps to set up an app registration for your Office Add-in.

1. Sign in to the [Azure portal](https://portal.azure.com/) with the ***admin*** credentials to your Microsoft 365 tenancy. For example, **MyName@contoso.onmicrosoft.com**.
1. Select **App registrations**. If you don't see the icon, search for "app registration" in the search bar.

    :::image type="content" source="../images/azure-portal-select-app-registration.png" alt-text="The Azure portal home page.":::

    The **App registrations** page appears.

1. Select **New registration**.

    :::image type="content" source="../images/azure-portal-select-new-registration.png" alt-text="New registration on the App registrations pane.":::

    The **Register an application** page appears.

1. On the **Register an application** page, set the values as follows.

    * Set **Name** to `contoso-office-add-in-sso`.
    * Set **Supported account types** to **Accounts in any organizational directory (any Azure AD directory - multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    * Set **Redirect URI** to use the platform **Single-page application (SPA)** and the URI to `brk-multihub://localhost:3000`. This redirect assumes you are testing your add-in from a localhost server.

    :::image type="content" source="../images/azure-portal-create-app-reg-naa.png" alt-text="Register an application pane with name and supported account completed.":::

1. Select **Register**. A message is displayed stating that the application registration was created.

    :::image type="content" source="../images/azure-portal-application-created-message.png" alt-text="Message stating that the application registration was created.":::

1. Copy and save the value for the **Application (client) ID**. You'll use it in a later procedure.

    :::image type="content" source="../images/azure-portal-copy-client-id.png" alt-text="App registration pane for Contoso displaying the client ID and directory ID.":::

Finally add an SPA redirect URI for your task pane page. This redirect URI is required when using MSAL with Word, Excel, and PowerPoint in the browser.

1. From the left pane, select **Manage > Authentication**.
    :::image type="content" source="../images/azure-portal-authentication-page.png" alt-text="The authentication page in the Azure app registration.":::
1. In the **Platform configurations** section there is a list of **Single-page application Redirect URIs**.
1. Select **Add URI**.
    :::image type="content" source="../images/azure-portal-add-uri.png" alt-text="Selecting the add URI option on the Azure app registration page.":::
1. Enter **https://localhost:3000/taskpane.html** and select **Save**. This redirect URI assume you are using NAA from the `taskpane.html` file.
    :::image type="content" source="../images/azure-portal-add-taskpane-redirect-uri.png" alt-text="Adding the taskpane redirect URI on the Azure app registration page.":::

## Add the MSAL library to your project

You need the MSAL library to support SSO in your project.

1. Add the `@azure/msal-browser` package to the `dependencies` section of the `package.json` file for your project. For more information on this package, see [Microsoft Authentication Library for JavaScript (MSAL.js) for Browser-Based Single-Page Applications](https://www.npmjs.com/package/%40azure/msal-browser). To install the latest version, run the following command.

    ```command&nbsp;line
    npm install @azure/msal-browser
    ```

1. Create a  new file named `auth.js` in your project. You'll put all the authentication and SSO related code in this file.
1. Add the following code to the top of the `auth.js` file. About the following code, note:
    - It defines a couple global variables that are used by following steps.
    - It creates a `setGlobalActiveAccount` helper function which is used by other functions to track the active account of the user correctly.

    ```javascript
    import { createNestablePublicClientApplication } from "@azure/msal-browser";
        
    let msalInstance = null;
    let currentAccount = null;
 
    /**
     * Helper function to set the global variable for the current account and the MSAL active account.
     * 
     * @param {*}  account The current account.
     */
    function setGlobalActiveAccount(account) {
        if (account) {
            currentAccount = account;
            msalInstance.setActiveAccount(currentAccount);
        }
    }
    ```

1. Save the file.

## Initialize the public client application

You need to initialize MSAL and get an instance of the [public client application](/entra/identity-platform/msal-client-applications). This is used to get access tokens when needed.

Add the following code to the `auth.js` file. About the following code note:

- It calls the`createNestablePublicClientApplication` function in MSAL. MSAL returns a public client application that can be nested in a native application host (for example, Outlook) to acquire tokens for your application.
- The `createNestablePublicClientApplication` requires an MSAL configuration. Replace the `Enter_client_ID_here` with the client ID from your app registration.

```javascript
/**
 * Initialize the MSAL PublicClientApplication instance.
 * This should be called once when the add-in loads.
 */
export async function initializeMsal() {
    if (msalInstance) return;

    try {
        msalInstance = await createNestablePublicClientApplication({
        auth: {
            clientId: "Enter_client_ID_here",
            authority: "https://login.microsoftonline.com/common",
        }});
        const response = await msalInstance.handleRedirectPromise();

        if (response?.account) {
            currentAccount = response.account;
        } else {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) currentAccount = accounts[0];
        }
    } catch (error) {
        console.error("MSAL initialization failed:", error);
        throw error;
    }
}
```

## Add a function to get the login hint

When running in Word, Excel, or PowerPoint in the browser, you must provide a login hint to MSAL to identify the correct account. You can get the login hint from the Office `authContext.userPrincipalName` property. Add the following function which can be called at any time to get the login hint.

Add the following function to the `auth.js` file.

```javascript
/**
 * Get the login hint from Office.AuthContext for a better SSO experience.
 * This is especially important for Office in a browser.
 */
async function getLoginHint() {
    try {
        if (typeof Office !== "undefined" && Office.context) {
            const authContext = await Office.auth.getAuthContext();
            if (authContext?.userPrincipalName) return authContext.userPrincipalName;
        }
    } catch (error) {
        console.warn("Could not get login hint:", error);
    }
    return null;
}

```

## Acquire the access token

Build a function to acquire the access token.  

Add the following function to the `auth.js` file. About the following code note:

- The function first checks if a current account is available for the signed in user. If a current account exists, then the function calls `acquireTokenSilent` to get an access token using that account.
- If there isn't a current account, or if `acquireTokenSilent` fails, the function calls `acquireTokenWithSSO` as the next step to get the access token.

```javascript
/**
 * Acquire an access token for a resource (such as Microsoft Graph). Attempts to acquire the token by using the current account.
 * Switches to SSO or interactive login if needed.
 * @param scopes - The scopes to request; for example, ["User.Read"].
 * @returns Access token string
 */
export async function acquireAccessToken(scopes) {
    if (!msalInstance) await initializeMsal();

    // Try to get the current account.
    try {
        if (!currentAccount) {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                setGlobalActiveAccount(accounts[0]);
            } else {
                // Acquire token using SSO if no account is available.
                return await acquireTokenWithSSO(scopes);
            }
        }

        // Try to acquire the token silently based on current account.
        const response = await msalInstance.acquireTokenSilent({
            scopes,
            account: currentAccount,
        });
        setGlobalActiveAccount(response.account);
        return response.accessToken;
    } catch (error) {
        console.error("Silent token acquisition failed:", error);
        // Fall back and acquire token using SSO.
        return await acquireTokenWithSSO(scopes);
    }
}
```

## Acquire the token with SSO

Build a function to acquire the access token through SSO. Add the following function to the `auth.js` file. About the following code note:

- The function creates a token request with the scopes and the login hint. Then it calls `ssoSilent` to attempt to silently get an access token.
- If the call fails, then  the function calls `acquireTokenInteractively` to interact with the user as the next step to get the access token.

```javascript
/**
 * Acquire an access token using SSO.
 * @param scopes - The scopes to request.
 * @returns Access token string.
 */
async function acquireTokenWithSSO(scopes) {
    try {
        // Create token request with login hint.
        const hint = await getLoginHint();
        const ssoTokenRequest = { scopes };
        if (hint) ssoTokenRequest.loginHint = hint;

        // Try to acquire token silently using SSO.
        const response = await msalInstance.ssoSilent(ssoTokenRequest);
        setGlobalActiveAccount(response.account);
        return response.accessToken;
    } catch (error) {
        console.error("SSO silent failed:", error);
        // Fall back to interactive token acquisition.
        return await acquireTokenInteractive(scopes);
    }
}
```

## Acquire the token interactively

Build a function to acquire the access token interactively.   

Add the following function to the `auth.js` file. About the following code note:

- The function creates a token request with the scopes and the login hint.
- Then it calls the MSAL API `acquireTokenPopup`. MSAL will show a popup window and interact with the user. The interaction could be to ask for consent to the scopes, perform multi-factor authentication, or to resolve a conditional access policy.
- The call should succeed unless the interaction can't be completed successfully (such as user canceling the popup).

```javascript
/**
 * Acquire an access token interactively via popup.
 * This is a fallback when silent authentication fails.
 * @param scopes - The scopes to request.
 * @returns Access token string.
 */
async function acquireTokenInteractive(scopes) {
    try {
        // Create token request with login hint.
        const hint = await getLoginHint();
        const popupTokenRequest = { scopes };
        if (hint) popupTokenRequest.loginHint = hint;

        // Use MSAL popup to acquire token interactively.
        const response = await msalInstance.acquireTokenPopup(popupTokenRequest);
        setGlobalActiveAccount(response.account);
        return response.accessToken;
    } catch (error) {
        console.error("Interactive token acquisition failed:", error);
        throw error;
    }
}

```

## Get token and call the Graph API

The next steps assume you have a `taskpane.js` file in a project built with `yo office` (**Office Add-in Task Pane** project). However these steps can be used in any project task pane.

In your `taskpane.js` file update the `run` function to get an access token and call the Graph API. Call the `acquireAccessToken` function you created previously and pass the `Files.Read` and `User.Read` scopes. Then use the access token to make a call the Microsoft Graph API to retrieve 10 of the user's file names from OneDrive.

Update the `run` function to match the following code.

```javascript
export async function run() {
   let accessToken;
  try {
    // Acquire an access token for Microsoft Graph.
    accessToken = await acquireAccessToken(["Files.Read","User.Read"]);
    console.log("Access token acquired:", accessToken);
  }
  catch (error) {
    console.error("Error acquiring access token:", error);
    return;
  }

  // Call the Microsoft Graph API with the access token.
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name&$top=10`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (response.ok) {
    // Write file names to the console.
    const data = await response.json();
    const names = data.value.map((item) => item.name);

    // Be sure the taskpane.html has an element with Id = item-subject.
    const label = document.getElementById("item-subject");

    // Write file names to task pane and the console.
    const nameText = names.join(", ");
    if (label) label.textContent = nameText;
    console.log(nameText);
  } else {
    const errorText = await response.text();
    console.error("Microsoft Graph call failed - error text: " + errorText);
  }
}
```

Once all the previous code is added to the `run` function, be sure a button on the task pane calls the `run` function. Then you can sideload the add-in and try out the code.

### Have a fallback when NAA isn't supported

While we strive to provide a high-degree of compatibility with these flows across the Microsoft ecosystem, your add-in may be loaded in an older Office host that does not support NAA. In these cases, your add-in won't support seamless SSO and you may need to fall back to an alternate method of authenticating the user. Refer to the [code samples](#code-samples) in this article for samples that show how to handle a fallback scenario.

Use the following code to check if NAA is supported when your add-in loads.

```javascript
   Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
```

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft.com](mailto:secure@microsoft.com) with as much detail as you can provide. Your submission may be eligible for a bounty through the [Microsoft Bounty](https://aka.ms/bugbounty) program. Don't post security issues to GitHub or any other public site. We'll contact you shortly after receiving your issue report. We encourage you to get new security incident notifications by visiting [Microsoft technical security notifications](https://technet.microsoft.com/security/dd252948) to subscribe to Security Advisory Alerts.

## Code samples

| Sample name | Description  |
| ----------- | ------------ |
| [Office Add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-SSO-NAA)   | Shows how to use NAA in an Office Add-in to access Microsoft Graph APIs for the signed-in user. |
| [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA) | Shows how to use NAA in an Outlook Add-in to access Microsoft Graph APIs for the signed-in user. |
|[Implement SSO in events in an Outlook add-in using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Event-SSO-NAA) |Shows how to use NAA and SSO in Outlook add-in events.|
|[Send identity claims to resources using nested app authentication (NAA) and SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-Identity)|Shows how to send the signed-in user's identity claims (such as name, email, or a unique ID) to a resource such as a database. This sample replaces an obsolete pattern for legacy Exchange Online tokens.|
|[Outlook add-in with SSO using nested app authentication including Internet Explorer fallback](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-IE)|Shows how to implement a fallback authentication strategy when NAA isn't available and the add-in needs to support [Outlook versions that still use Internet Explorer 11](../concepts/browsers-used-by-office-web-add-ins.md).|

## See also

- [NAA FAQ](https://aka.ms/NAAFAQ)
- [Nested app authentication in Microsoft Teams](/microsoftteams/platform/concepts/authentication/nested-authentication).
- [Outlook sample: How to fall back and support Internet Explorer 11](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Outlook-Add-in-SSO-NAA-IE/README.md)
- [Authenticate and authorize with the Office dialog API](/office/dev/add-ins/develop/auth-with-office-dialog-api).
- [Microsoft identity sample for SPA and JavaScript](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md)
- [Microsoft identity samples for various app types and frameworks](/entra/identity-platform/sample-v2-code?tabs=apptype)