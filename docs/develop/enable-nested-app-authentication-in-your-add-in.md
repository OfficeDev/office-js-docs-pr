---
title: Enable single sign-on in an Office Add-in with nested app authentication
description: Learn how to enable SSO in an Office Add-in with nested app authentication.
ms.date: 06/24/2025
ms.topic: how-to
ms.localizationpriority: high
---

# Enable single sign-on in an Office Add-in with nested app authentication

You can use the MSAL.js library with nested app authentication to use single sign-on (SSO) from your Office Add-in. Using nested app authentication (NAA) offers several advantages over the On-Behalf-Of (OBO) flow.

- You only need to use the MSAL.js library and don’t need the `getAccessToken` function in Office.js.
- You can call services such as Microsoft Graph with an access token from your client code as an SPA. There’s no need for a middle-tier server.
- You can use incremental and dynamic consent for scopes.
- You don't need to [preauthorize your hosts](/microsoftteams/platform/m365-apps/extend-m365-teams-personal-tab?tabs=manifest-teams-toolkit#update-azure-ad-app-registration-for-sso) (for example, Teams, Office) to call your endpoints.

## NAA supported accounts and hosts

NAA supports both Microsoft Accounts and Microsoft Entra ID (work/school) identities. It doesn't support [Azure Active Directory B2C](/azure/active-directory-b2c/overview) for business-to-consumer identity management scenarios. The following table explains the current support by platform. Platforms listed as generally available (GA) are ready for production usage in your add-in.

| Application | Web        | Windows                                              | Mac        | iOS/iPad           | Android        |
|-------------|------------|------------------------------------------------------|------------|--------------------|----------------|
| Excel       | In preview | In preview                                           | In preview | In preview on iPad | Not applicable |
| Outlook     | GA         | GA in Current Channel and Monthly Enterprise Channel, Preview in Semi-Annual Channels | GA         | GA (iOS)           | GA             |
| PowerPoint  | In preview | In preview                                           | In preview | In preview on iPad | Not applicable |
| Word        | In preview | In preview                                           | In preview | In preview on iPad | Not applicable |

> [!IMPORTANT]
> To use NAA on platforms that are still in preview (Word, Excel, and PowerPoint), join the [Microsoft 365 Insider Program](https://aka.ms/MSFT365InsiderProgram) and choose **Current Channel (Preview)**. Don't use NAA in production add-ins for any preview platforms. We invite you to try out NAA in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

## Register your single-page application

You’ll need to create a Microsoft Azure App registration for your add-in on the Azure portal. The app registration must have at minimum:

- A name
- A supported account type
- An SPA redirect

If your add-in requires additional app registration beyond NAA and SSO, see [Single-page application: App registration](/entra/identity-platform/scenario-spa-app-registration).

## Add a trusted broker through SPA redirect

To enable NAA, your app registration must include a specific redirect URI to indicate to the Microsoft identity platform that your add-in allows itself to be brokered by supported hosts. The redirect URI of the application must be of type **Single Page Application** and conform to the following scheme.

`brk-multihub://your-add-in-domain`

Your domain must include only the origin and not its subpaths. For example:

✔️ brk-multihub://localhost:3000<br>
✔️ brk-multihub://www.contoso.com<br>
❌ brk-multihub://www.contoso.com/go

Trusted broker groups are dynamic by design and can be updated in the future to include additional hosts where your add-in may use NAA flows. Currently the brk-multihub group includes Office Word, Excel, PowerPoint, Outlook, and Teams (for when Office is activated inside).

## Configure MSAL config to use NAA

Configure your add-in to use NAA by calling the `createNestablePublicClientApplication` function in MSAL. MSAL returns a public client application that can be nested in a native application host (for example, Outlook) to acquire tokens for your application.

The following steps show how to enable NAA in the `taskpane.js` or `taskpane.ts` file in a project built with `yo office` (**Office Add-in Task Pane** project).

1. Add the `@azure/msal-browser` package to the `dependencies` section of the `package.json` file for your project. For more information on this package, see [Microsoft Authentication Library for JavaScript (MSAL.js) for Browser-Based Single-Page Applications](https://www.npmjs.com/package/%40azure/msal-browser). We recommend using the latest version of the package (at time of the last article update it was 3.26.0).

    ```json
    "dependencies": {
        "@azure/msal-browser": "^3.27.0",
        ...
    ```

1. Save and run `npm install` to install `@azure/msal-browser`.

1. Add the following code to the top of the `taskpane.js` or `taskpane.ts` file. This will import the MSAL browser library.

    ```JavaScript
    import { createNestablePublicClientApplication } from "@azure/msal-browser";
    ```

## Initialize the public client application

Next, you need to initialize MSAL and get an instance of the [public client application](/entra/identity-platform/msal-client-applications). This is used to get access tokens when needed. We recommend that you put the code that creates the public client application in the `Office.onReady` method.

- In your `Office.onReady` function, add a call to `createNestablePublicClientApplication` as shown below. Replace the `Enter_the_Application_Id_Here` placeholder with the Azure app ID you saved previously.

    ```javascript
    let pca = undefined;
    Office.onReady(async (info) => {
      if (info.host) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
  
        // Initialize the public client application
        pca = await createNestablePublicClientApplication({
          auth: {
            clientId: "Enter_the_Application_Id_Here",
            authority: "https://login.microsoftonline.com/common"
          },
        });
      }
    });
    ```

> [!NOTE]
> The previous code sample sets the **authority** to **common**, which supports work and school accounts or personal Microsoft accounts. If you want to configure a single tenant or other account types, see [Application configuration options](/entra/identity-platform/msal-client-application-configuration) for additional authority options.

## Acquire your first token

The tokens acquired by MSAL.js via NAA will be issued for your Azure app registration ID. In this code sample, you acquire a token for the Microsoft Graph API. If the user has an active session with Microsoft Entra ID the token is acquired silently. If not, the library prompts the user to sign in interactively. The token is then used to call the Microsoft Graph API.

The following steps show the pattern to use for acquiring a token.

1. Specify your scopes. NAA supports incremental and dynamic consent so always request the minimum scopes needed for your code to complete its task.
1. Call `acquireTokenSilent`. This will get the token without requiring user interaction.
1. If `acquireTokenSilent` fails, call `acquireTokenPopup` to display an interactive dialog for the user. `acquireTokenSilent` can fail if the token expired, or the user has not yet consented to all of the requested scopes.

The following code shows how to implement this authentication pattern in your own project.

1. Replace the `run` function in `taskpane.js` or `taskpane.ts` with the following code. The code specifies the minimum scopes needed to read the user's files.

    ```javascript
    async function run() {
    // Specify minimum scopes needed for the access token.
    const tokenRequest = {
      scopes: ["Files.Read", "User.Read", "openid", "profile"],
    };
    let accessToken = null;

    // TODO 1: Call acquireTokenSilent.

    // TODO 2: Call acquireTokenPopup.
    
    // TODO 3: Log error if token still null.
    
    // TODO 4: Call the Microsoft Graph API.

    }
    ```

    > [!IMPORTANT]
    > The token request must include scopes other than just `offline_access`, `openid`, `profile`, or `email`. You can use any combination of the previous scopes, but you must include at least one additional scope. If not, the token request can fail.

1. Replace `TODO 1` with the following code. This code calls `acquireTokenSilent` to get the access token.

    ```JavaScript
    try {
      console.log("Trying to acquire token silently...");
      const userAccount = await pca.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      accessToken = userAccount.accessToken;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }
    ```

1. Replace `TODO 2` with the following code. This code checks if the access token is acquired. If not it attempts to interactively get the access token by calling `acquireTokenPopup`.

    ```javascript
    if (accessToken === null) {
      // Acquire token silent failure. Send an interactive request via popup.
      try {
        console.log("Trying to acquire token interactively...");
        const userAccount = await pca.acquireTokenPopup(tokenRequest);
        console.log("Acquired token interactively.");
        accessToken = userAccount.accessToken;
      } catch (popupError) {
        // Acquire token interactive failure.
        console.log(`Unable to acquire token interactively: ${popupError}`);
      }
    }
    ```

1. Replace `TODO 3` with the following code. If both silent and interactive sign-in failed, log the error and return.

    ```javascript
    // Log error if both silent and popup requests failed.
    if (accessToken === null) {
      console.error(`Unable to acquire access token.`);
      return;
    }
    ```

## Call an API

After acquiring the token, use it to call an API. The following example shows how to call the Microsoft Graph API by calling `fetch` with the token attached in the *Authorization* header.

- Replace `TODO 4` with the following code.

    ```javascript
    // Call the Microsoft Graph API with the access token.
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name&$top=10`,
      {
        headers: { Authorization: accessToken },
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
    ```

Once all the previous code is added to the `run` function, be sure a button on the task pane calls the `run` function. Then you can sideload the add-in and try out the code.

## What is nested app authentication

Nested app authentication enables SSO for applications that are nested inside of supported Microsoft applications. For example, Excel on Windows runs your add-in inside a webview. In this scenario, your add-in is a nested application running inside Excel, which is the host. NAA also supports nested apps in Teams. For example, if a Teams tab is hosting Excel, and your add-in is loaded, it is nested inside Excel, which is also nested inside Teams. Again, NAA supports this nested scenario and you can access SSO to get user identity and access tokens of the signed in user.

## Best practices

We recommend the following best practices when using MSAL.js with NAA.

### Use silent authentication whenever possible

MSAL.js provides the `acquireTokenSilent` method that handles token renewal by making silent token requests without prompting the user. The method first looks for a valid cached token. If it doesn't find one, the library makes the silent request to Microsoft Entra ID and if there's an active user session, a fresh token is returned.

In certain cases, the `acquireTokenSilent` method's attempt to get the token fails. Some examples of this are when there's an expired user session with Microsoft Entra ID or a password change by the user, which requires user interaction. When the acquireTokenSilent fails, you need to call the interactive `acquireTokenPopup` token method.

### Have a fallback when NAA isn't supported

While we strive to provide a high-degree of compatibility with these flows across the Microsoft ecosystem, your add-in may be loaded in an older Office host that does not support NAA. In these cases, your add-in won't support seamless SSO and you may need to fall back to an alternate method of authenticating the user. In general you'll want to use the MSAL SPA authentication pattern with the [Office JS dialog API](auth-with-office-dialog-api.md).

Use the following code to check if NAA is supported when your add-in loads.

```javascript
   Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
```

For more information, see the following resources.

- [Outlook sample: How to fall back and support Internet Explorer 11](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Outlook-Add-in-SSO-NAA-IE/README.md)
- [Authenticate and authorize with the Office dialog API](/office/dev/add-ins/develop/auth-with-office-dialog-api).
- [Microsoft identity sample for SPA and JavaScript](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md)
- [Microsoft identity samples for various app types and frameworks](/entra/identity-platform/sample-v2-code?tabs=apptype)

## MSAL.js APIs supported by NAA

The following table shows which APIs are supported when NAA is enabled in the MSAL config.

| Method                        | Supported by NAA      |
| ----------------------------- | --------------------- |
| `acquireTokenByCode`          | No (throws exception) |
| `acquireTokenPopup`           | Yes                   |
| `acquireTokenRedirect`        | No (throws exception) |
| `acquireTokenSilent`          | Yes                   |
| `addEventCallback`            | Yes                   |
| `addPerformanceCallback`      | No (throws exception) |
| `disableAccountStorageEvents` | No (throws exception) |
| `enableAccountStorageEvents`  | No (throws exception) |
| `getAccountByHomeId`          | Yes                   |
| `getAccountByLocalId`         | Yes                   |
| `getAccountByUsername`        | Yes                   |
| `getActiveAccount`            | Yes                   |
| `getAllAccounts`              | Yes                   |
| `getConfiguration`            | Yes                   |
| `getLogger`                   | Yes                   |
| `getTokenCache`               | No (throws exception) |
| `handleRedirectPromise`       | No                    |
| `initialize`                  | Yes                   |
| `initializeWrapperLibrary`    | Yes                   |
| `loginPopup`                  | Yes                   |
| `loginRedirect`               | No (throws exception) |
| `logout`                      | No (throws exception) |
| `logoutPopup`                 | No (throws exception) |
| `logoutRedirect`              | No (throws exception) |
| `removeEventCallback`         | Yes                   |
| `removePerformanceCallback`   | No (throws exception) |
| `setActiveAccount`            | No                    |
| `setLogger`                   | Yes                   |
| `ssoSilent`                   | Yes                   |

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft.com](mailto:secure@microsoft.com) with as much detail as you can provide. Your submission may be eligible for a bounty through the [Microsoft Bounty](https://aka.ms/bugbounty) program. Don't post security issues to GitHub or any other public site. We'll contact you shortly after receiving your issue report. We encourage you to get new security incident notifications by visiting [Microsoft technical security notifications](https://technet.microsoft.com/security/dd252948) to subscribe to Security Advisory Alerts.

## Code samples

| Sample name | Description  |
| ----------- | ------------ |
| [Office Add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-SSO-NAA)   | Shows how to use MSAL.js nested app authentication (NAA) in an Office Add-in to access Microsoft Graph APIs for the signed-in user. The sample displays the signed-in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into the document.        |
| [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA) | Shows how to use MSAL.js nested app authentication (NAA) in an Outlook Add-in to access Microsoft Graph APIs for the signed-in user. The sample displays the signed-in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into a new message body. |

## See also

- [NAA FAQ](https://aka.ms/NAAFAQ)
