---
title: Create a Node.js Office Add-in that uses single sign-on
description: Learn how to create a Node.js-based add-in that uses Office Single Sign-on.
ms.date: 05/20/2023
ms.localizationpriority: medium
---

# Create a Node.js Office Add-in that uses single sign-on

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

This article walks you through the process of enabling single sign-on (SSO) in an add-in. The sample add-in you create has two parts; a task pane that loads in Microsoft Excel, and a middle-tier server that handles calls to Microsoft Graph for the task pane. The middle-tier server is built with Node.js and Express and exposes a single REST API, `/getuserfilenames`, that returns a list of the first 10 file names in the user's OneDrive folder. The task pane uses the `getAccessToken()` method to get an access token for the signed in user to the middle-tier server. The middle-tier server uses the On-Behalf-Of flow (OBO) to exchange the access token for a new one with access to Microsoft Graph. You can extend this pattern to access any Microsoft Graph data. The task pane always calls a middle-tier REST API (passing the access token) when it needs Microsoft Graph services. The middle-tier uses the token obtained via OBO to call Microsoft Graph services and return the results to the task pane.

This article works with an add-in that uses Node.js and Express. For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).

## Prerequisites

- [Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/en/about/previous-releases) version)

- [Git Bash](https://git-scm.com/downloads) (or another git client)

- A code editor - we recommend Visual Studio Code

- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription

- A build of Microsoft 365 that supports the [IdentityAPI 1.3 requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets). You might qualify for a Microsoft 365 E5 developer subscription, which includes a developer sandbox, through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). The [developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) includes a Microsoft Azure subscription that you can use for app registrations in later steps in this article. If you prefer, you can use a separate Microsoft Azure subscription for app registrations. Get a trial subscription at [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Set up the starter project

1. Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

   > [!NOTE]
   > There are two versions of the sample:
   >
   > - The **Begin** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
   > - The **Complete** folder contains the same sample with all coding steps from this article completed. To use the completed version, just follow the instructions in this article, but replace "Begin" with "Complete" and skip the sections **Code the client side** and **Code the middle-tier server** side.

1. Open a command prompt in the **Begin** folder.

1. Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.

1. Run the command `npm run install-dev-certs`. Select **Yes** to the prompt to install the certificate.

Use the following values for placeholders for the subsequent app registration steps.

| Placeholder           | Value                                 |
|-----------------------|---------------------------------------|
| `<add-in-name>`       | **Office-Add-in-NodeJS-SSO**          |
| `<fully-qualified-domain-name>` | `localhost:3000` |
| Microsoft Graph permissions | profile, openid, Files.Read |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## Configure the add-in

1. Open the `\Begin` folder in the cloned project in your code editor.

1. Open the `.ENV` file and use the values that you copied earlier from the **Office-Add-in-NodeJS-SSO** app registration. Set the values as follows:

    | Name              | Value                                                            |
    | ----------------- | ---------------------------------------------------------------- |
    | **CLIENT_ID**     | **Application (client) ID** from app registration overview page. |
    | **CLIENT_SECRET** | **Client secret** saved from **Certificates & Secrets** page.    |

    The values should **not** be in quotation marks. When you are done, the file should be similar to the following:

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    SERVER_SOURCE=<https://localhost:3000>
    ```

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:3000/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. Replace the placeholder "$app-id-guid$" _in both places_ in the markup with the **Application ID** that you copied when you created the **Office-Add-in-NodeJS-SSO** app registration. The "$" symbols are not part of the ID, so don't include them. This is the same ID you used for the CLIENT_ID in the .ENV file.

   > [!NOTE]
   > The `<Resource>` value is the **Application ID URI** you set when you registered the add-in. The `<Scopes>` section is used only to generate a consent dialog box if the add-in is sold through Microsoft Marketplace.

1. Open the `\public\javascripts\fallback-msal\authConfig.js` file. Replace the placeholder "$app-id-guid$" with the application ID that you saved from the **Office-Add-in-NodeJS-SSO** app registration you created previously.

1. Save the changes to the file.

## Code the client-side

### Call our web server REST API

1. In your code editor, open the file `public\javascripts\ssoAuthES6.js`. It already has code that ensures that Promises are supported, even in the Trident (Internet Explorer 11) webview control, and an `Office.onReady` call to assign a handler to the add-in's only button.

   > [!NOTE]
   > As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API. When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will support Trident.

1. In the `getFileNameList` function, replace `TODO 1` with the following code. About this code, note:

    - The function `getFileNameList` is called when the user chooses the **Get OneDrive File Names** button on the task pane.
    - It calls the `callWebServerAPI` function specifying which REST API to call. This returns JSON containing a list of file names from the user's OneDrive.
    - The JSON is passed to the `writeFileNamesToOfficeDocument` function to list the file names in the document.

    ```javascript
    try {
        const jsonResponse = await callWebServerAPI('GET', '/getuserfilenames');
        if (jsonResponse === null) {
            // Null is returned when a message was displayed to the user
            // regarding an authentication error that cannot be resolved.
            return;
        }
        await writeFileNamesToOfficeDocument(jsonResponse);
        showMessage('Your OneDrive filenames are added to the document.');
    } catch (error) {
        console.log(error.message);
        showMessage(error.message);
    }
    ```

1. In the `callWebServerAPI` function, replace `TODO 2` with the following code. About this code, note:

    - The function calls `getAccessToken` which is our own function that encapsulates using Office SSO or MSAL fallback as necessary to get the token. If it returns a null token, a message was shown for an auth error condition that cannot be resolved, so the function also returns null.
    - The function uses the `fetch` API to call the web server and if successful, returns the JSON body.

    ```javascript
    const accessToken = await getAccessToken(authSSO);
    if (accessToken === null) {
        return null;
    }
    const response = await fetch(path, {
        method: method,
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken,
        },
    });

    // Check for success condition: HTTP status code 2xx.
    if (response.ok) {
        return response.json();
    }
    ```

1. In the `callWebServerAPI` function, replace `TODO 3` with the following code. About this code, note:

    - This code handles the scenario where the SSO token expired. If so we need to call `Office.auth.getAccessToken` to get a refreshed token. The simplest way is to make a recursive call which results in a new call to `Office.auth.getAccessToken`. The `retryRequest` parameter ensures the recursive call is only attempted once.
    - The `TokenExpiredError` string is set by our web server whenever it detects an expired token.

    ```javascript
     // Check for fail condition: Is SSO token expired? If so, retry the call which will get a refreshed token.
    const jsonBody = await response.json();
    if (
        authSSO === true &&
        jsonBody != null &&
        jsonBody.type === 'TokenExpiredError'
    ) {
        if (!retryRequest) {
            return callWebServerAPI(method, path, true); // Try the call again. The underlying call to Office JS getAccessToken will refresh the token.
        } else {
            // Indicates a second call to retry and refresh the token failed.
            authSSO = false;
            return callWebServerAPI(method, path, true); // Try the call again, but now using MSAL fallback auth.
        }
    }
    ```

1. In the `callWebServerAPI` function, replace `TODO 4` with the following code. About this code, note:

    - The `Microsoft Graph` string is set by our web server whenever a Microsoft Graph call fails.

    ```javascript
    // Check for fail condition: Did we get a Microsoft Graph API error, which is returned as bad request (403)?
    if (response.status === 403 && jsonBody.type === 'Microsoft Graph') {
        throw new Error('Microsoft Graph error: ' + jsonBody.errorDetails);
    }
    ```

1. In the `callWebServerAPI` function, replace `TODO 5` with the following code.

    ```javascript
    // Handle other errors.
    throw new Error(
        'Unknown error from web server: ' + JSON.stringify(jsonBody)
    );
    ```

1. In the `getAccessToken` function, replace `TODO 6` with the following code. About this code, note:

    - `authSSO` tracks if we are using SSO, or using MSAL fallback. If SSO is used, the function calls `Office.auth.getAccessToken` and returns the token.
    - Errors are handled by the `handleSSOErrors` function which will return a token if it switches to fallback MSAL authentication.
    - Fallback authentication uses the MSAL library to sign in the user. The add-in itself is an SPA, and uses an SPA app registration to access the web server.

    ```javascript
    if (authSSO) {
        try {
            // Get the access token from Office host using SSO.
            // Note that Office.auth.getAccessToken modifies the options parameter. Create a copy of the object
            // to avoid modifying the original object.
            const options = JSON.parse(JSON.stringify(ssoOptions));
            const token = await Office.auth.getAccessToken(options);
            return token;
        } catch (error) {
            console.log(error.message);
            return handleSSOErrors(error);
        }
    } else {
        // Get access token through MSAL fallback.
        try {
            const accessToken = await getAccessTokenMSAL();
            return accessToken;
        } catch (error) {
            console.log(error);
            throw new Error(
                'Cannot get access token. Both SSO and fallback auth failed. ' +
                    error
            );
        }
    }
    ```

1. In the `handleSSOErrors` function, replace `TODO 7` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    switch (error.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one
            // is logged into Office, then the first call of getAccessToken should pass the
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see
            // this error.
            showMessage(
                'No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.'
            );
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage(
                'You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.'
            );
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage(
                'Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.'
            );
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage(
                'Office is still working on the last operation. When it completes, try this operation again.'
            );
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage(
                "Follow the instructions to change your browser's zone configuration."
            );
            break;
    ```

1. Replace `TODO 8` with the following code. For any errors that can't be handled the code switches to fallback authentication using MSAL.

    ```javascript
    default: //recursive call.
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to MSAL sign-in.
            showMessage('SSO failed. Trying fallback auth.');
            authSSO = false;
            return getAccessToken(false);
    }
    return null; // Return null for errors that show a message to the user.
    ```

## Code the web server REST API

The web server provides REST APIs for the client to call. For example, the REST API `/getuserfilenames` gets a list of filenames from the user's OneDrive folder. Each REST API call requires an access token by the client to ensure the correct client is accessing their data. The access token is exchanged for a Microsoft Graph token through the On-Behalf-Of flow (OBO). The new Microsoft Graph token is cached by the MSAL library for subsequent API calls. It's never sent outside of the web server. For more information, see [Middle-tier access token request](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### Create the route and implement On-Behalf-Of flow

1. Open the file `routes\getFilesRoute.js` and replace `TODO 9` with the following code. About this code, note:

    - It calls `authHelper.validateJwt`. This ensures the access token is valid and hasn't been tampered with.
    - For more information, see [Validating tokens](/azure/active-directory/develop/access-tokens#validating-tokens).

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 10: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. Replace `TODO 10` with the following code. About this code, note:

    - It only requests the minimum scopes it needs, such as `files.read`.
    - It uses the MSAL `authHelper` to perform the OBO flow in the call to `acquireTokenOnBehalfOf`.

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(' ')[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(' ');
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === 'access_as_user'
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 11: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 12: Handle any errors.
    }
    ```

1. Replace `TODO 11` with the following code. About this code, note:

    - It constructs the URL for the Microsoft Graph API call and then makes the call via the `getGraphData` function.
    - It returns errors by sending an HTTP 500 response along with details.
    - On success it returns the JSON with the filename list to the client.

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = '/me/drive/root/children';

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = '?$select=name&$top=10';

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    // TODO 12: Check for expired token.
    ```

1. Replace `TODO 12` with the following code. This code specifically checks if the token expired because the client can request a new token and call again.

    ```javascript
    } catch (err) {
       // On rare occasions the SSO access token is unexpired when Office validates it,
       // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
       // with "The provided value for the 'assertion' is not valid. The assertion has expired."
       // Construct an error message to return to the client so it can refresh the SSO token.
       if (err.errorMessage.indexOf('AADSTS500133') !== -1) {
         res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
       } else {
         res.status(403).send({ type: "Unknown", errorDetails: err });
       }
    }
    ```

The sample must handle both fallback authentication through MSAL and SSO authentication through Office. The sample will try SSO first, and the `authSSO` boolean at the top of the file tracks if the sample is using SSO or has switched to fallback auth.

## Run the project

1. Ensure that you have some files in your OneDrive so that you can verify the results.

1. Open a command prompt in the root of the `\Begin` folder.

1. Run the command `npm install` to install all package dependencies.

1. Run the command `npm start` to start the middle-tier server.

1. You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it. The instructions depend on your platform. There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.

1. Click the **Get OneDrive File Names** button. If you're logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected the first 10 file and folder names in your OneDrive for Business are inserted into the document. (It may take as much as 15 seconds the first time.) If you're not logged in, or you're in a scenario that doesn't support SSO, or SSO isn't working for any reason, you'll be prompted to sign in. After you sign in, the file and folder names appear.

> [!NOTE]
> If you were previously signed into Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to _close all other Office applications_ before you press **Get OneDrive File Names**.

## Stop running the project

When you're ready to stop the middle-tier server and uninstall the add-in, follow these instructions:

1. Run the following command to stop the middle-tier server.

    ```command&nbsp;line
    npm stop
    ```

1. To uninstall or remove the add-in, see the specific sideload article you used for details.

## Security notes

- The `/getuserfilenames` route in `getFilesroute.js` uses a literal string to compose the call for Microsoft Graph. If you change the call so that any part of the string comes from user input, sanitize the input so that it cannot be used in a Response header injection attack.

- In `app.js` the following content security policy is in place for scripts. You may want to specify additional restrictions depending on your add-in security needs.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Always follow security best practices in the [Microsoft identity platform documentation](/azure/active-directory/develop/).
