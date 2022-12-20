---
title: Create a Node.js Office Add-in that uses single sign-on
description: Learn how to create a Node.js-based add-in that uses Office Single Sign-on.
ms.date: 12/20/2022
ms.localizationpriority: medium
---

# Create a Node.js Office Add-in that uses single sign-on

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

This article walks you through the process of enabling single sign-on (SSO) in an add-in. The sample add-in you create has two parts; a task pane that loads in Microsoft Excel, and a middle-tier server that handles calls to Microsoft Graph for the task pane. The middle-tier server is built with Node.js and Express and exposes a single REST API, `/getuserfilenames`, that returns a list of the first 10 file names in the user's OneDrive folder. The task pane uses the `getAccessToken()` method to get an access token for the signed in user to the middle-tier server. The middle-tier server uses the On-Behalf-Of flow (OBO) to exchange the access token for a new one with access to Microsoft Graph. You can extend this pattern to access any Microsoft Graph data. The task pane always calls a middle-tier REST API (passing the access token) when it needs Microsoft Graph services. The middle-tier uses the token obtained via OBO to call Microsoft Graph services and return the results to the task pane.

This article works with an add-in that uses Node.js and Express. For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).

## Prerequisites

- [Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)

- [Git Bash](https://git-scm.com/downloads) (or another git client)

- A code editor - we recommend Visual Studio Code

- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription

- A build of Microsoft 365 that supports the [IdentityAPI 1.3 requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets). You can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription. The developer sandbox includes a Microsoft Azure subscription that you can use for app registrations in later steps in this article. If you prefer, you can use a separate Microsoft Azure subscription for app registrations. Get a trial subscription at [Microsoft Azure](https://account.windowsazure.com/SignUp).

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
| `<fully-qualified-domain-name>` | `localhost:44355` |
| Microsoft Graph permissions | profile, openid, Files.Read |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## Configure the add-in

1. Open the `\Begin` folder in the cloned project in your code editor.

1. Open the `.ENV` file and use the values that you copied earlier from the **Office-Add-in-NodeJS-SSO** app registration. Set the values as follows:

   | Name              | Value                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | **Application (client) ID** from app registration overview page. |
   | **CLIENT_SECRET** | **Client secret** saved from **Certificates & Secrets** page.       |
   | **DIRECTORY_ID**  | **Directory (tenant) ID** from app registration overview page.   |

   The values should **not** be in quotation marks. When you are done, the file should be similar to the following:

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. Replace the placeholder "$app-id-guid$" _in both places_ in the markup with the **Application ID** that you copied when you created the **Office-Add-in-NodeJS-SSO** app registration. The "$" symbols are not part of the ID, so don't include them. This is the same ID you used for the CLIENT_ID in the .ENV file.

   > [!NOTE]
   > The **\<Resource\>** value is the **Application ID URI** you set when you registered the add-in. The **\<Scopes\>** section is used only to generate a consent dialog box if the add-in is sold through AppSource.

1. Open the `\public\javascripts\fallback-msal\authConfig.js` file. Replace the placeholder "$app-id-guid$" with the application ID that you saved from the **Office-Add-in-NodeJS-SSO** app registration you created previously.

1. Save the changes to the file.

## Code the client-side

### Create client request and response handler

1. In your code editor, open the file `public\javascripts\ssoAuthES6.js`. It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.

   > [!NOTE]
   > As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API. When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will support Internet Explorer 11.

    A key part of the sample code is the client request. The client request is an object that tracks information about the request for calling REST APIs on the middle-tier server. It's necessary because the client request state needs to be tracked or updated through the following scenario:

    - SSO fails and fallback authentication is required. The access token is acquired through MSAL in a pop-up dialog box. The goal is to not fail in this scenario and gracefully fall back to the alternative authentication approach.

    The client request object tracks the following data:

    - `authSSO` - true if using SSO, otherwise false.
    - `verb` - REST API verb such as GET and POST.
    - `accessToken` - The access token to the ASP.NET Core server.
    - `url` - The URL of the REST API to call on the ASP.NET Core server.
    - `callbackRESTApiHandler` - The function to pass the results of the REST API call.
    - `callbackFunction` - the function to pass the client request to when ready.

1. To initialize the client request object, in the `createRequest` function, replace `TODO 1` with the following code.

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. Replace `TODO 2` with the following code. About this code, note:

    - It checks if SSO is being used. The method to acquire the access token is different for SSO than for fallback auth.
    - If SSO returns the access token, it calls the `callbackfunction` function. For fallback authentication it calls `dialogFallback`, which will eventually call the callback function after the user signs in through MSAL.

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. In the `getFileNameList` function, replace `TODO 3` with the following code. About this code, note:

    - The function `getFileNameList` is called when the user chooses the **Get OneDrive File Names** button on the task pane.
    - It creates a client request to track information about the call, such as the URL of the REST API.
    - When the REST API returns a result, it's passed to the `handleGetFileNameResponse` function. This callback is passed as a parameter to `createRequest` and is tracked in `clientRequest.callbackRESTApiHandler`.
    - The code calls `callWebServer` with the client request to perform next steps and call the REST API.

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. In the `handleGetFileNameResponse` function, replace `TODO 4` with the following code. About this code, note:

    - The code passes the response (which contains a list of filenames) to `writeFileNamesToOfficeDocument` to write the filenames to the document.
    - The code checks for errors. It shows a success message if the filenames are written, otherwise it shows an error.

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. In the `handleSSOErrors` function, replace `TODO 5` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. Replace `TODO 6` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). For any errors that can't be handled, `true` is returned to the caller. This indicates the caller should switch to using MSAL as fallback auth.

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### Call the REST API on the middle-tier server

1. In the `callWebServer` function, replace `TODO 7` with the following code. About this code, note:

    - The actual AJAX call will be made by the `ajaxCallToRESTApi` function.
    - This function will attempt to get a new access token if the middle-tier server returns an error indicating that the current token expired.
    - If the AJAX call cannot be completed successfully, `switchToFallbackAuth` will be called to use MSAL authentication instead of Office SSO.

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. Replace `TODO 8` with the following code. About this code, note:

    - When the server identifies an expired token, it returns an error with the type "TokenExpiredError".
    - The try...catch will call Office.auth.getAccessToken to get a refreshed token with a new expiration.
    - The code will attempt to call the server API again.

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. Replace `TODO 9` with the following code. About this code, note:

    - For **Microsoft Graph** errors, show the message on the task pane.
    - For all other messages, show the message on the task pane.

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

Fallback authentication uses the MSAL library to sign in the user. The add-in itself is an SPA, and uses an SPA app registration to access the middle-tier server.

1. In the `switchToFallbackAuth` function, replace `TODO 10` with the following code. About this code, note:

    - It sets the global `authSSO` to false and creates a new client request that uses MSAL for auth. The new request has an MSAL access token to the middle-tier server.
    - Once the request is created it calls `callWebServer` to continue attempting to call the middle-tier server successfully.

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## Code the middle-tier server

The middle-tier server provides REST APIs for the client to call. For example, the REST API `/getuserfilenames` gets a list of filenames from the user's OneDrive folder. Each REST API call requires an access token by the client to ensure the correct client is accessing their data. The access token is exchanged for a Microsoft Graph token through the On-Behalf-Of flow (OBO). The new Microsoft Graph token is cached by the MSAL library for subsequent API calls. It's never sent outside of the middle-tier server. For more information, see [Middle-tier access token request](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### Create the route and implement On-Behalf-Of flow

1. Open the file `routes\getFilesRoute.js` and replace `TODO 11` with the following code. About this code, note:

    - It calls `authHelper.validateJwt`. This ensures the access token is valid and hasn't been tampered with.
    - For more information, see [Validating tokens](/azure/active-directory/develop/access-tokens#validating-tokens).

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. Replace `TODO 12` with the following code. About this code, note:

    - It only requests the minimum scopes it needs, such as `files.read`.
    - It uses the MSAL `authHelper` to perform the OBO flow in the call to `acquireTokenOnBehalfOf`.

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. Replace `TODO 13` with the following code. About this code, note:

    - It constructs the URL for the Microsoft Graph API call and then makes the call via the `getGraphData` function.
    - It returns errors by sending an HTTP 500 response along with details.
    - On success it returns the JSON with the filename list to the client.

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

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
    ```

1. Replace `TODO 14` with the following code. This code specifically checks if the token expired because the client can request a new token and call again.

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
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

## Security notes

- The `/getuserfilenames` route in `getFilesroute.js` uses a literal string to compose the call for Microsoft Graph. If you change the call so that any part of the string comes from user input, sanitize the input so that it cannot be used in a Response header injection attack.

- In `app.js` the following content security policy is in place for scripts. You may want to specify additional restrictions depending on your add-in security needs.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Always follow security best practices in the [Microsoft identity platform documentation](/azure/active-directory/develop/).
