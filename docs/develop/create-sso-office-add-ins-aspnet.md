---
title: Create an ASP.NET Office Add-in that uses single sign-on
description: A step-by-step guide for how to create (or convert) an Office Add-in with an ASP.NET backend to use single sign-on (SSO).
ms.date: 07/29/2022
ms.localizationpriority: medium
---

# Create an ASP.NET Office Add-in that uses single sign-on

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. This article walks you through the process of enabling single sign-on (SSO) in an add-in.

The sample shows you how to build the following parts:

- Client-side code that provides a task pane that loads in Microsoft Excel. The client-side code calls the Office JS API `getAccessToken()` to get the SSO access token to call server-side REST APIs. If this fails, it will fallback and use the Microsoft Authentication library [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) to obtain the access token.
- Server-side code that uses ASP.NET Core to provide a single REST API `/api/filenames`. The server-side code uses [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet) for all token handling, authentication, and authorization.

The sample uses SSO and the On-Behalf-Of (OBO) flow to obtain correct access tokens and call Microsoft Graph APIs. If you are unfamiliar with how this flow works, see [How SSO works at runtime](authorize-to-microsoft-graph.md#how-it-works-at-runtime) for more detail.

## Prerequisites

- Visual Studio 2019 or later.

- The **Office/SharePoint development** workload when configuring Visual studio.

- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.

- A build of Microsoft 365 that supports the [IdentityAPI 1.3 requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets). You can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription. The developer sandbox includes a Microsoft Azure subscription that you can use for app registrations in later steps in this article. If you prefer, you can use a separate Microsoft Azure subscription for app registrations. Get a trial subscription at [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Set up the starter project

Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

> [!NOTE]
> There are two versions of the sample.
>
> - The **Begin** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
> - The **Complete** folder contains the same sample with all coding steps from this article completed. To use the completed version, just follow the instructions in this article, but replace "Begin" with "Complete" and skip the sections **Code the client side** and **Code the server side**.

## Register the add-in with Microsoft identity platform

You need to create an app registration in Azure that represents your ASP.NET Core server. This enables authentication support so that proper access tokens can be issued to the client code in JavaScript. The app registration is used for SSO requests from Office, and also fallback authentication if SSO fails.

1. To register your app, navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.

1. Sign in with the **_admin_** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

   - Set **Name** to `Office-Add-in-ASPNET-SSO`.
   - Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox).**.
   - In the **Redirect URI** section, set the platform to **Single-page application (SPA)** with a redirect URI value of `https://localhost:7283/dialog.html`.
   - Choose **Register**.

   > [!NOTE]
   > The SPA application type is only used when the client uses MSAL.js for fallback authentication.

1. On the **Office-Add-in-ASPNET-SSO** page, copy and save the **Application (client) ID**. You'll use it in later procedures.

   > [!NOTE]
   > This **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It's also the "client ID" of the application when it seeks authorized access to Microsoft Graph.

1. In the leftmost sidebar, select **Authentication** under **Manage**. In the **Implicit grant and hybrid flows** section, select both checkboxes for **Access tokens** and **ID tokens**.

1. Choose **Save**.

1. Under **Manage**, select **Certificates & secrets** and select **New client secret**. Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.

   The web application uses the client secret **Value** to prove its identity when it requests tokens. _Record this value for use in a later step - it's shown only once._

1. In the leftmost sidebar, select **Expose an API** under **Manage**. Select the **Set** link. This will generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.

1. In the generated ID, insert `localhost:7283/` (note the forward slash "/" appended to the end) between the double forward slashes and the GUID. When you are finished, the entire ID should have the form `api://localhost:7283/$App ID GUID$`; for example `api://localhost:7283/c6c1f32b-5e55-4997-881a-753cc1d563b7`. Then choose **Save**.

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope name**.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user. Suggestions:

   - **Admin consent display name**: Office can act as the user.
   - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.
   - **User consent display name**: Office can act as you.
   - **User consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope**.

   > [!NOTE]
   > The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Authorized client applications** section, select **Add a client application** button and then, in the panel that opens, set the Client ID to `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`, and then select the **Authorized scopes** checkbox for `api://localhost:7283/$app-id-guid$/access_as_user`.

1. Select **Add application**.

   > [!NOTE]
   > The `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID pre-authorizes all Microsoft Office application endpoints. It's also required if you want to support Microsoft accounts (MSA) on Office on Windows and Mac. Alternatively, you can enter a proper subset of the following IDs if for any reason you want to deny authorization to Office on some platforms. Just leave out the IDs of the platforms from which you want to withhold authorization. Users of your add-in on those platforms will not be able to call your Web APIs, but other functionality in your add-in will still work.
   >
   > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
   > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
   > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. In the leftmost sidebar, select **API permissions** under **Manage** and select **Add a permission**. On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search for the permissions your add-in needs. Select the following. Only the first is really required by your add-in itself; but the `profile` and `openid` permissions are required for the Office application to get an access token with user identity to access the ASP.NET Core server.

   - **Files.Read**
   - **profile**
   - **openid**

   > [!NOTE]
   > The `User.Read` permission may already be listed by default. It's a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in doesn't actually need it.

1. Select the check box for each permission as it appears. After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.

1. On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.

## Configure the solution

1. In the root of the **Begin** folder, open the solution (.sln) file in **Visual Studio**. Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.

1. Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**. Ensure that the **Action** for both projects is set to **Start**, and that the **Office-Add-in-ASPNET-SSO-web** project is listed first. Close the dialog.

1. In **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO-manifest** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file. Just above the end `</VersionOverrides>` tag, you'll find the following markup.

    ```xml
    <WebApplicationInfo>
	    <Id>$app-id-guid$</Id>
		<Resource>api://localhost:7283/$app-id-guid$</Resource>
		<Scopes>
            <Scope>Files.Read</Scope>
			<Scope>profile</Scope>
            <Scope>openid</Scope>
		</Scopes>
	</WebApplicationInfo>
    ```

1. Replace the placeholder "$app-id-guid$" _in both places_ in the markup with the **Application ID** that you copied when you created the **Office-Add-in-ASPNET-SSO** app registration. The "$" symbols are not part of the ID, so don't include them. This is the same ID you used for the application ID in the appsettings.json file.

   > [!NOTE]
   > The **\<Resource\>** value is the **Application ID URI** you set when you registered the add-in. The **\<Scopes\>** section is used only to generate a consent dialog box if the add-in is sold through AppSource.

1. Save and close the manifest file.

1. In **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO-web** project and open the **appsettings.json** file.

1. Replace the placeholder **$app-id-guid$** with the **Application (client) ID** value you saved previously.

1. Replace the placeholder **$client-secret$** with the client secret value you saved previously.

1. Replace the **$publisher-domain$** with the domain name of your app registration. You can find your publisher domain by going to the App registration you created previously. In the leftmost sidebar, select **Manifest** under **Manage**. Search for the **publisherDomain** value in the manifest JSON.

    > [!NOTE]
    > You can also change the **TenantId** to support single-tenant if you configured your app registration for single-tenant. Replace the **Common** value with the **Application (client) ID** for single-tenant support.

1. Save and close the appsettings.json file.

1. Open the **wwwroot/js/fallback-msal/authConfig.js** file. This file specifies the configuration for MSAL.js to use for fallback authentication.

1. Replace the placeholder **$app-id-guid$** with the **Application (client) ID** value you saved previously.

1. Save and close the **authConfig.js** file.

## Code the client side

### Create client request and response handler

1. In the **Office-Add-in-ASPNET-SSO-web** project, open the **wwwroot\js\ssoAuthES6.js** file.  It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.

   > [!NOTE]
   > As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API. When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will support Internet Explorer 11.

    A key part of the sample code is the client request. The client request is an object that tracks information about the request for calling REST APIs on the server. It's necessary because the client request state needs to be tracked or updated through the following scenarios:

    - SSO retries where the REST API call fails because it needs additional consent. The sample code calls `getAccessToken` with updated authentication options, gets the required user consent, and then calls the REST API again. The goal is to not fail in scenarios where a REST API needs additional consent.
    - SSO fails and fallback authentication is required. The access token is acquired through MSAL.js in a pop-up dialog box. The goal is to not fail in this scenario and gracefully fall back to the alternative authentication approach.

    The client request object tracks the following data:

    - `authOptions` - [Auth configuration parameters](/javascript/api/office/office.authoptions) for SSO.
    - `authSSO` - true if using SSO, otherwise false.
    - `verb` - REST API verb such as GET, POST...
    - `accessToken` - The access token to the ASP.NET Core server. The method to obtain this token is different for SSO than fallback auth.
    - `url` - The URL of the REST API to call on the ASP.NET Core server.
    - `callbackHandler` - The function to pass the results of the REST API call.
    - `callbackFunction` - The function to pass the client request to when ready.

1. To initialize the client request object, in the `createRequest` function, replace `TODO 1` with the following code.

    ```javascript
    const clientRequest = {
        authOptions: {
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true,
        },
        authSSO: authSSO,
        verb: verb,
        accessToken: null,
        url: url,
        callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. Replace `TODO 2` with the following code. About this code, note:

   - It checks if SSO is being used by checking `authSSO`. If so, it calls `getAccessToken` to get the SSO access token.
   - If SSO returns the access token, it calls the `callbackfunction` function. However, if `getAccessToken` failed, it handles the error through the `handleSSOErrors()` function and switches to fallback if required.
   - If `authSSO` is false, then fallback auth is in use, so it calls the `dialogFallback` function, which will eventually call the callback function after the user signs in through MSAL.js.

   ```javascript
   if (authSSO) {
        try {
            // Get access token from Office SSO.
            clientRequest.accessToken = await Office.auth.getAccessToken(clientRequest.authOptions);
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
        "/api/filenames",
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

### Get SSO access token and call REST API

1. In the `callWebServer` function, replace `TODO 5` with the following code. About this code, note:

    - It issues an AJAX call based on settings in the `clientRequest` object.

    ```javascript
    try {
        const data = await $.ajax({
            type: clientRequest.verb,
            url: clientRequest.url,
            headers: { Authorization: "Bearer " + clientRequest.accessToken },
            cache: false
        });
        clientRequest.callbackRESTApiHandler(data);
    } catch (error) {
        // TODO 6: Handle any token or Microsoft Graph API errors.

    }
    ```

1. In the `callWebServer` function, replace `TODO 6` with the following code. About this code, note:

    - In the rare case the original SSO token is expired, it will detect this error condition and request a refreshed token by calling `getAccessToken` again. It will retry the call, but if there is still an error it switches to fallback auth.
    - It checks if the call from the server to Microsoft Graph failed, in which case the error status is a bad request (403). In this case it displays the error, but there is no need to switch to fallback auth.
    - For all other errors it displays them in the task pane.

    ```javascript
    // Check for expired token. Refresh and retry the call if it expired.
        if (error.getResponseHeader !== undefined) {
            const responseHeader = error.getResponseHeader("www-authenticate");
            if (responseHeader !== null && responseHeader.includes("The token expired") && authSSO) {
                try {
                    clientRequest.accessToken = await Office.auth.getAccessToken(clientRequest.authOptions);
                    const data = await $.ajax({
                        type: clientRequest.verb,
                        url: clientRequest.url,
                        headers: { Authorization: "Bearer " + clientRequest.accessToken },
                        cache: false
                    });
                    clientRequest.callbackRESTApiHandler(data);
                } catch (error) {
                    showMessage(error.responseText);
                    switchToFallbackAuth(clientRequest);
                    return;
                }
            }
        }
    }
    ```

1. Replace `TODO 6` with the following:

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. Replace `TODO 7` with the following. Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange. Azure AD will respond with error AADSTS500133. When this happens, the code  recalls the SSO API (but no more than once). This time Office returns a new unexpired bootstrap token.

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Replace `TODO 8` with the following:

    ```javascript
    else {
        dialogFallback();
    }
    ```

        // Check for a Microsoft Graph API call error. which is returned as bad request (403)
        if (error.status === 403) {
            showMessage(error.reponseText);
            return;
        }

        // For all other error scenarios, display the message and use fallback auth.
        showMessage(
            "Unknown error from web server: " +
            JSON.stringify(error.responseJSON.errorDetails)
        );
        if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

1. In the `handleSSOErrors` function, replace `TODO 7` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).

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

1. Replace `TODO 8` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). For any errors that can't be handled, `true` is returned to the caller. This indicates the caller should switch to using MSAL.js as fallback auth.

   ```javascript
    default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            fallbackRequired = true;
            break;
    }
    return fallbackRequired;
   ```

### Add fallback to .js authentication

Fallback authentication uses the [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) library to sign in the user. The add-in itself is an SPA, and uses an SPA app registration to access the ASP.NET Core server.

1. In the `switchToFallbackAuth` function, replace `TODO 9` with the following code. About this code, note:

   - It sets the global `authSSO` to false and creates a new client request that uses MSAL.js for authentication. The new request has an MSAL.js access token to the server.
   - Once the request is created it calls `callWebServer` to continue attempting to call the server successfully.

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

## Code the server side

The server-side code is an ASP.NET Core server that provides REST APIs for the client to call. For example, the REST API `/api/filenames` gets a list of filenames from the user's OneDrive folder. Each REST API call requires an access token by the client to ensure the correct client is accessing their data. The access token is exchanged for a Microsoft Graph token through the On-Behalf-Of flow (OBO). The new Microsoft Graph token is cached by the MSAL.NET library for subsequent API calls. It's never sent outside of the server-side code. Microsoft identity documentation refers to this server as the middle-tier server because it is in the middle of the flow from client-side code to Microsoft services. For more information, see [Middle-tier access token request](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### Configure Microsoft Graph and OBO flow

1. Open the `Program.cs` file and replace `TODO 10` with the following code. About this code, note:

    - It adds required services to handle token validation that is required for the REST APIs.
    - It adds Microsoft Graph and OBO flow support in the call to `EnableTokenAcquisitionToCallDownstreamApi().AddMicrosoftGraph(...)`. The OBO flow is handled automatically for you, and the Microsoft Graph SDK is provided to your REST API controllers.
    - The **AzureAd** and **DownstreamApi** configurations are specified in the **appsettings.json** file.

    ```csharp
    // Add services to the container.
    builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
        .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));
    
    var config = builder.Configuration;
    
    builder.Services.AddMicrosoftIdentityWebApiAuthentication(config)
                        .EnableTokenAcquisitionToCallDownstreamApi()
                            .AddMicrosoftGraph(config.GetSection("DownstreamApi"))
                            .AddInMemoryTokenCaches();
    ```

### Create the /api/filenames REST API

1. In the **Controllers** folder, open the **FileNamesController.cs** file. replace `TODO 11` with the following code. About this code, note:

    - It specifies the `[Authorize]` attribute to ensure the access token is validated for each call to any REST APIs in the `FileNamesController` class. For more information, see [Validating tokens](/azure/active-directory/develop/access-tokens#validating-tokens).
    - It specifies the `[RequiredScope("access_as_user")]` attribute to ensure the client has the correct **access_as_user** scope in the access token.
    - The constructor initializes the `_graphServiceClient` object to make calling Microsoft Graph REST APIs easier.

    ```csharp
    [Authorize]
    [Route("api/[controller]")]
    [RequiredScope("access_as_user")]
    public class FileNamesController : Controller
    {
        public FileNamesController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphServiceClient, IOptions<MicrosoftGraphOptions> graphOptions)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphServiceClient = graphServiceClient;
            _graphOptions = graphOptions;
        }
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly GraphServiceClient _graphServiceClient;
        private readonly IOptions<MicrosoftGraphOptions> _graphOptions;

    // TODO 12: Add the REST API for getting filenames.

    }
    ```

1. Replace `TODO 12` with the following code. About this code, note:

    - It creates the `/api/filenames` REST API.
    - It has error handlers, including for MSAL.NET exceptions.

    ```csharp
    [HttpGet]
        public async Task<IActionResult> Get()
        {
            try
            {
                // Get list of first 10 file names from user's OneDrive root folder.
                var filelist = await _graphServiceClient.Me.Drive.Root.Children.Request().Select(u => new
                {
                    u.Name
                }).Top(10).GetAsync();

               // Map result to just the file names.
               List<string> files = new List<string>();
                foreach (var file in filelist)
                {
                    files.Add(file.Name);
                }

                return Ok(files);
            }
            catch (MsalException ex)
            {
                return StatusCode((int)HttpStatusCode.Unauthorized, "An authentication error occurred while acquiring a token for the Microsoft Graph API.\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException challengeException)
                {
                    _tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(_graphOptions.Value.Scopes.Split(' '),
                        challengeException.MsalUiRequiredException);
                }
                else
                {
                    return StatusCode((int)HttpStatusCode.BadRequest, "An error occurred while calling the Microsoft Graph API.\n" + ex.Message);
                }
            }
            
            return StatusCode((int)HttpStatusCode.InternalServerError);
        }
    ```

## Run the solution

1. Open the Visual Studio solution file.
1. On the **Build** menu, select **Clean Solution**. When it finishes, open the **Build** menu again and select **Build Solution**.
1. In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO-manifest** project node.
1. In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).

    :::image type="content" source="../images/select-host.png" alt-text="Choose the desired Office client application: Excel, PowerPoint, or Word.":::

1. Press F5.
1. In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.
1. Click the **Get OneDrive File Names** button. If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane. If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to sign in. After you sign in, the file and folder names appear.

### Testing the fallback path

To test the fallback authorization path, force the SSO path to fail with the following steps.

1. Open the **wwwroot/js/ssoAuthES6.js** file and change the line `let authSSO=true` to `let authSSO=false`.

## Deploy the add-in

When you are ready to deploy to a staging or production server, be sure to update the following areas in the project solution.

- In the **appsettings.json** file, change the **domain** to your staging or production domain name.
- Update any references to `localhost:7283` throughout your project to use your staging or production URL.
- Update any references to `localhost:7283` in your Azure App registration, or create a new registration for use in staging or production.

For more information, see [Host and deploy ASP.NET Core](/aspnet/core/host-and-deploy/).

## See also

- [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).
- [Authorize to Microsoft Graph with SSO](authorize-to-microsoft-graph.md).
