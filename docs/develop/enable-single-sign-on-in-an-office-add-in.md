---
title: Enable single sign-on in an Office Add-in with nested app authentication
description: Learn how to enable SSO in an Office Add-in with nested app authentication.
ms.date: 11/24/2025
ms.topic: how-to
ms.localizationpriority: high
---

# Enable single sign-on in an Office Add-in with nested app authentication

Office add-in

The following steps in this article show how to enable NAA in the `taskpane.js` or `taskpane.ts` file in a project built with `yo office` (**Office Add-in Task Pane** project).

## Add the MSAL library to your project

You need the MSAL library to support SSO in your project.

1. Add the `@azure/msal-browser` package to the `dependencies` section of the `package.json` file for your project. For more information on this package, see [Microsoft Authentication Library for JavaScript (MSAL.js) for Browser-Based Single-Page Applications](https://www.npmjs.com/package/%40azure/msal-browser). To install the latest version, run the following command.

    ```command&nbsp;line
    npm install @azure/msal-browser
    ```

1. Create a  new file named `auth.js` or `auth.ts` in your project. You'll put all the authentication and SSO related code in this file.
1. Add the following code to the top of the `auth.js` or `auth.ts` file. This will import the MSAL browser library and initialize a few global variables.

    ```TypeScript
    import { createNestablePublicClientApplication } from "@azure/msal-browser";
    
    const MSAL_CONFIG = {
        auth: {
            clientId: "a0189bd8-06c8-4063-9d6e-c66977f03693",
            authority: "https://login.microsoftonline.com/common",
        },
    };
        
    let msalInstance = null;
    let currentAccount = null;
    let loginHint = null;
    ```

1. Save the file.

## Initialize the public client application

You need to initialize MSAL and get an instance of the [public client application](/entra/identity-platform/msal-client-applications). This is used to get access tokens when needed.

Add the following code to the `auth.js` or `auth.ts` file.

```typescript
/**
 * Initialize the MSAL PublicClientApplication instance.
 * This should be called once when the add-in loads.
 */
export async function initializeMsal() {
    if (msalInstance) return;

    try {
        msalInstance = await createNestablePublicClientApplication(MSAL_CONFIG);
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

Add the following function to the `auth.js` or `auth.ts` file.

```javascript
/**
 * Get the login hint from Office.AuthContext for a better SSO experience.
 * This is especially important for Office in a browser.
 */
async function getLoginHint() {
    if (loginHint) return loginHint;

    try {
        if (typeof Office !== "undefined" && Office.context) {
            const authContext = await Office.auth.getAuthContext();
            if (authContext?.userPrincipalName) {
                loginHint = authContext.userPrincipalName;
                return loginHint;
            }
        }
    } catch (error) {
        console.warn("Could not get login hint:", error);
    }

    return null;
}

```

## Acquire the access token

Build a function to acquire the access token. The function first checks if a current account is available for the signed in user. If a current account exists, then the function calls `acquireTokenSilent` to get an access token using that account. If there isn't a current account, or if `acquireTokenSilent` fails, the function calls `acquireTokenWithSSO` as the next step to get the access token.

Add the following function to the `auth.js` or `auth.ts` file.

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
                currentAccount = accounts[0];
                msalInstance.setActiveAccount(currentAccount);
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

        if (response.account) {
            currentAccount = response.account;
            msalInstance.setActiveAccount(currentAccount);
        }

        return response.accessToken;
    } catch (error) {
        console.error("Silent token acquisition failed:", error);
        // Fall back and acquire token using SSO.
        return await acquireTokenWithSSO(scopes);
    }
}
```

## Acquire the token with SSO

Build a function to acquire the access token through SSO. The function creates a token request with the scopes and the login hint. Then it calls `ssoSilent` to attempt to silently get an access token. If the call fails, then  the function calls `acquireTokenInteractively` to interact with the user as the next step to get the access token.

Add the following function to the `auth.js` or `auth.ts` file.

```javascript
**
 * Acquire an access token using SSO.
 * @param scopes - The scopes to request.
 * @returns Access token string.
 */
async function acquireTokenWithSSO(scopes) {
    if (!msalInstance) await initializeMsal();

    try {
        // Create token request with scopes and the login hint.
        const hint = await getLoginHint();
        const ssoRequest = { scopes };
        if (hint) ssoRequest.loginHint = hint;

        // Try to acquire token silently using SSO.
        const response = await msalInstance.ssoSilent(ssoRequest);

        if (response.account) {
            currentAccount = response.account;
            msalInstance.setActiveAccount(currentAccount);
        }

        return response.accessToken;
    } catch (error) {
        console.error("SSO silent failed:", error);
        // Fall back to interactive token acquisition.
        return await acquireTokenInteractive(scopes);
    }
}
```

## Acquire the token interactively

Build a function to acquire the access token interactively. The function creates a token request with the scopes and the login hint. Then it calls the MSAL API `acquireTokenPopup`. MSAL will show a popup window and interact with the user. This could just be to ask for consent to the scopes, perform Multi-factor authentication, or complete other interactions such as compliance with a conditional access policy. The call will fail if the user doesn't complete required steps, or can't resolve the conditional access policy requirements.

Add the following function to the `auth.js` or `auth.ts` file.

```javascript

/**
 * Acquire an access token interactively (popup or redirect).
 * This is a fallback when silent authentication fails.
 * @param scopes - The scopes to request.
 * @returns Access token string.
 */
async function acquireTokenInteractive(scopes) {
    if (!msalInstance) await initializeMsal();

    try {
        // Create token request with scopes and the login hint.
        const hint = await getLoginHint();
        const popupRequest = { scopes };
        if (hint) popupRequest.loginHint = hint;

        const response = await msalInstance.acquireTokenPopup(popupRequest);

        if (response.account) {
            currentAccount = response.account;
            msalInstance.setActiveAccount(currentAccount);
        }

        return response.accessToken;
    } catch (error) {
        console.error("Interactive token acquisition failed:", error);
        throw error;
    }
}

```

## Get token and call the Graph API

In your `taskpane.js` or `taskpane.ts` file update the `run` function to get an access token and call the Graph API. Call the `acquireAccessToken` function you created previously and pass the `Files.Read` and `User.Read` scopes. Then user the access token to make a call the Microsoft Graph API to retrieve 10 of the user's file names from OneDrive.

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

