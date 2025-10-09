---
title: Create an ASP.NET Office Add-in that uses single sign-on (SSO)
description: A step-by-step guide for how to create (or convert) an Office Add-in with an ASP.NET backend to use single sign-on (SSO).
ms.date: 05/20/2023
ms.localizationpriority: medium
---

# Create an ASP.NET Office Add-in that uses single sign-on

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. This article walks you through the process of enabling single sign-on (SSO) in an add-in.

The sample shows you how to build the following parts:

- Client-side code that provides a task pane that loads in Microsoft Excel, Word, or PowerPoint. The client-side code calls the Office JS API `getAccessToken()` to get the SSO access token to call server-side REST APIs.
- Server-side code that uses ASP.NET Core to provide a single REST API `/api/files`. The server-side code uses the [Microsoft Authentication Library for .NET (MSAL.NET)](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet) for all token handling, authentication, and authorization.

The sample uses SSO and the On-Behalf-Of (OBO) flow to obtain correct access tokens and call Microsoft Graph APIs. If you are unfamiliar with how this flow works, see [How SSO works at runtime](authorize-to-microsoft-graph.md#how-it-works-at-runtime) for more detail.

## Prerequisites

- Visual Studio 2019 or later.

- The **Office/SharePoint development** workload when configuring Visual Studio.

- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.

- A build of Microsoft 365 that supports the [IdentityAPI 1.3 requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets). You might qualify for a Microsoft 365 E5 developer subscription, which includes a developer sandbox, through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). The [developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) includes a Microsoft Azure subscription that you can use for app registrations in later steps in this article. If you prefer, you can use a separate Microsoft Azure subscription for app registrations. Get a trial subscription at [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Set up the starter project

Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

> [!NOTE]
> There are two versions of the sample.
>
> - The **Begin** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
> - The **Complete** folder contains the same sample with all coding steps from this article completed. To use the completed version, just follow the instructions in this article, but replace "Begin" with "Complete" and skip the sections **Code the client side** and **Code the server side**.

Use the following values for placeholders for the subsequent app registration steps.

| Placeholder           | Value                                           |
|-----------------------|-------------------------------------------------|
| `<add-in-name>`       | **Office-Add-in-ASPNET-SSO**                    |
| `<fully-qualified-domain-name>` | `localhost:44355` |
| Microsoft Graph permissions | profile, openid, Files.Read |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## Configure the solution

1. In the root of the **Begin** folder, open the solution (.sln) file in **Visual Studio**. Right-click (or select and hold) the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.

1. Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**. Ensure that the **Action** for both projects is set to **Start**, and that the **Office-Add-in-ASPNETCoreWebAPI** project is listed first. Close the dialog.

1. In **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO-manifest** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file. Just above the end `</VersionOverrides>` tag, you'll find the following markup.

    ```xml
    <WebApplicationInfo>
	    <Id>Enter_client_ID_here</Id>
		<Resource>api://localhost:44355/Enter_client_ID_here</Resource>
		<Scopes>
            <Scope>Files.Read</Scope>
			<Scope>profile</Scope>
            <Scope>openid</Scope>
		</Scopes>
	</WebApplicationInfo>
    ```

1. Replace the placeholder "Enter_client_ID_here" _in both places_ in the markup with the **Application ID** that you copied when you created the **Office-Add-in-ASPNET-SSO** app registration. This is the same ID you used for the application ID in the appsettings.json file.

   > [!NOTE]
   > The `<Resource>` value is the **Application ID URI** you set when you registered the add-in. The `<Scopes>` section is used only to generate a consent dialog box if the add-in is sold through Microsoft Marketplace.

1. Save and close the manifest file.

1. In **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO-web** project and open the **appsettings.json** file.

1. Replace the placeholder **Enter_client_id_here** with the **Application (client) ID** value you saved previously.

1. Replace the placeholder **Enter_client_secret_here** with the client secret value you saved previously.

    > [!NOTE]
    > You must also change the **TenantId** to support single-tenant if you configured your app registration for single-tenant. Replace the **Common** value with the **Application (client) ID** for single-tenant support.

1. Save and close the appsettings.json file.

## Code the client side

### Get the access token and call the application server REST API

1. In the **Office-Add-in-ASPNETCore-WebAPI** project, open the **wwwroot\js\HomeES6.js** file.  It already has code that ensures that Promises are supported, even in the Trident (Internet Explorer 11) webview control, and an `Office.onReady` call to assign a handler to the add-in's only button.

   > [!NOTE]
   > As the name suggests, the HomeES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API. When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will support Trident.

1. In the `getUserFileNames` function, replace `TODO 1` with the following code. About this code, note:

    - It calls `Office.auth.getAccessToken` to get the access token from Office using SSO. This token will contain the user's identity as well as access permission to the application server.
    - The access token is passed to `callRESTApi` which makes the actual call to the application server. The application server then uses the OBO flow to call Microsoft Graph.
    - Any errors from calling `getAccessToken` will be handled by `handleClientSideErrors`.

    ```javascript
       let fileNameList = null;
    try {
        let accessToken = await Office.auth.getAccessToken(options);
        fileNameList = await callRESTApi("/api/files", accessToken);
    }
    catch (exception) {
        if (exception.code) {
            handleClientSideErrors(exception);
        }
        else {
            showMessage("EXCEPTION: " + exception);
        }
    }

    ```

1. In the `getUserFileNames` function, replace `TODO 2` with the following code. This will write the list of file names to the document.

   ```javascript
    try {
        await writeFileNamesToOfficeDocument(fileNameList);
        showMessage("Your data has been added to the document.");
    } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin 
        // "Unable to add filenames to document."
        showMessage(error);
    }
   ```

1. In the `callRESTApi` function, replace `TODO 3` with the following code. About this code, note:

   - It constructs an authorization header containing the access token. This confirms to the application server that this client code has access permissions to the REST APIs.
   - It request JSON return types, so that all return values are handled in JSON.
   - Any errors are passed to `handleServerSideErrors` for processing.

   ```javascript
    try {
        let result = await $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
            dataType: "json",
            contentType: "application/json; charset=utf-8"
        });
        return result;
    } catch (error) {
        handleServerSideErrors(error);
    }
   ```

### Handle SSO errors and application REST API errors

1. In the `handleSSOErrors` function, replace `TODO 4` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).

   ```javascript
    switch (error.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option.
            showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.");
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.");
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage("Office is still working on the last operation. When it completes, try this operation again.");
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage("Follow the instructions to change your browser's zone configuration.");
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in by using MSAL authentication.
            showMessage("SSO failed. In these cases you should implement a falback to MSAL authentication.");
            break;
    }
   ```

1. In the `handleServerSideErrors` function, replace `TODO 5` with the following code.

    ```javascript
    // Check headers to see if admin has not consented.
    const header = errorResponse.getResponseHeader('WWW-Authenticate');
    if (header !== null && header.includes('proposedAction=\"consent\"')) {
        showMessage("MSAL ERROR: " + "Admin consent required. Be sure admin consent is granted on all scopes in the Azure app registration.");
        return;
    }

    ```

1. In the `handleServerSideErrors` function, replace `TODO 6` with the following code. About this code, note:

    - In some cases, additional consent is required, such as 2FA. Microsoft identity returns the additional claims that are required to complete consent. This code adds the `authChallenge` property with the additional claims and calls `getUserfileNames` again. When `getAccessToken` is called again with the additional claims, the user gets a prompt for all required forms of authentication.

    ```javascript
    // Check if Microsoft Graph requires an additional form of authentication. Have the Office host 
    // get a new token using the Claims string, which tells Microsoft identity to prompt the user for all 
    // required forms of authentication.
    const errorDetails = JSON.parse(errorResponse.responseJSON.value.details);
    if (errorDetails) {
        if (errorDetails.error.message.includes("AADSTS50076")) {
            const claims = errorDetails.message.Claims;
            const claimsAsString = JSON.stringify(claims);
            getUserFileNames({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. In the `handleServerSideErrors` function, replace `TODO 7` with the following code. About this code, note:

    - In the rare case the original SSO token is expired, it will detect this error condition and call `getUserFilenames` again. This results in another call to `getAccessToken` which returns a refreshed access token. The `retryGetAccessToken` variable counts the retries and is currently configured to only retry once.
    - Finally, if an error cannot be handled, the default is to display the error in the task pane.

    ```javascript
    // Results from other errors (other than AADSTS50076) will have an ExceptionMessage property.
    const exceptionMessage = JSON.parse(errorResponse.responseText).ExceptionMessage;
    if (exceptionMessage) {
        // On rare occasions the access token is unexpired when Office validates it,
        // but expires by the time it is sent to Microsoft identity in the OBO flow. Microsoft identity will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Retry the call of getAccessToken (no more than once). This time Office will return a 
        // new unexpired access token.
        if ((exceptionMessage.includes("AADSTS500133"))
            && (retryGetAccessToken <= 0)) {
            retryGetAccessToken++;
            getUserFileNames();
            return;
        }
        else {
            showMessage("MSAL error from application server: " + JSON.stringify(exceptionMessage));
            return;
        }
    }
    // Default error handling if previous checks didn't apply.
    showMessage(errorResponse.responseJSON.value);
    ```

1. Save the file.

## Code the server side

The server-side code is an ASP.NET Core server that provides REST APIs for the client to call. For example, the REST API `/api/files` gets a list of filenames from the user's OneDrive folder. Each REST API call requires an access token by the client to ensure the correct client is accessing their data. The access token is exchanged for a Microsoft Graph token through the On-Behalf-Of flow (OBO). The new Microsoft Graph token is cached by the MSAL.NET library for subsequent API calls. It's never sent outside of the server-side code. Microsoft identity documentation refers to this server as the middle-tier server because it is in the middle of the flow from client-side code to Microsoft services. For more information, see [Middle-tier access token request](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### Configure Microsoft Graph and OBO flow

1. Open the `Program.cs` file and replace `TODO 8` with the following code. About this code, note:

    - It adds required services to handle token validation that is required for the REST APIs.
    - It adds Microsoft Graph and OBO flow support in the call to `EnableTokenAcquisitionToCallDownstreamApi().AddMicrosoftGraph(...)`. The OBO flow is handled automatically for you, and the Microsoft Graph SDK is provided to your REST API controllers.
    - The **DownstreamApi** configuration is specified in the **appsettings.json** file.

    ```csharp
    // Add services to the container.
    builder.Services.AddMicrosoftIdentityWebApiAuthentication(builder.Configuration)
                    .EnableTokenAcquisitionToCallDownstreamApi()
                        .AddMicrosoftGraph(builder.Configuration.GetSection("DownstreamApi"))
                        .AddInMemoryTokenCaches();

    ```

### Create the /api/filenames REST API

1. In the **Controllers** folder, open the **FilesController.cs** file. replace `TODO 9` with the following code. About this code, note:

    - It specifies the `[Authorize]` attribute to ensure the access token is validated for each call to any REST APIs in the `FilesController` class. For more information, see [Validating tokens](/azure/active-directory/develop/access-tokens#validating-tokens).
    - It specifies the `[RequiredScope("access_as_user")]` attribute to ensure the client has the correct **access_as_user** scope in the access token.
    - The constructor initializes the `_graphServiceClient` object to make calling Microsoft Graph REST APIs easier.

    ```csharp
    [Authorize]
    [Route("api/[controller]")]
    [RequiredScope("access_as_user")]
    public class FilesController : Controller
    {        
        public FilesController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphServiceClient, IOptions<MicrosoftGraphOptions> graphOptions)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphServiceClient = graphServiceClient;
            _graphOptions = graphOptions;

        }

        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly GraphServiceClient _graphServiceClient;
        private readonly IOptions<MicrosoftGraphOptions> _graphOptions;

        // TODO 10: Add the REST API to get filenames.

    }
    ```

1. Replace `TODO 10` with the following code. About this code, note:

    - It creates the `/api/files` REST API.
    - It handles exceptions from MSAL through `MsalException` class.
    - It handles exceptions from Microsoft Graph API calls through the `ServiceException` class.

    ```csharp
     // GET api/files
        [HttpGet]
        [Produces("application/json")]
        public async Task<IActionResult> Get()
        {
            List<DriveItem> result = new List<DriveItem>();
            try
            {
                var files = await _graphServiceClient.Me.Drive.Root.Children.Request()
                    .Top(10)
                    .Select(m => new { m.Name })
                    .GetAsync();

                result = files.ToList();
            }
            catch (MsalException ex)
            {
                var errorResponse = new
                {
                    message = "An authentication error occurred while acquiring a token for downstream API",
                    details = ex.Message
                };

                return StatusCode((int)HttpStatusCode.Unauthorized, Json(errorResponse));
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException challengeException)
                {
                    _tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(_graphOptions.Value.Scopes.Split(' '),
                        challengeException.MsalUiRequiredException);
                }
                else
                {
                    var errorResponse = new
                    {
                        message = "An error occurred calling Microsoft Graph",
                        details = ex.RawResponseBody
                    };
                    return StatusCode((int)HttpStatusCode.BadRequest, Json(errorResponse));
                }
            }
            catch (Exception ex)
            {
                var errorResponse = new
                {
                    message = "An error occurred while calling the downstream API",
                    details = ex.Message
                };
                return StatusCode((int)HttpStatusCode.BadRequest, Json(errorResponse));

            }
            return Json(result);
        }
    ```

## Run the solution

1. In Visual Studio, on the **Build** menu, select **Clean Solution**. When it finishes, open the **Build** menu again and select **Build Solution**.
1. In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO-manifest** project node.
1. In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).

    :::image type="content" source="../images/select-host.png" alt-text="Choose the desired Office client application: Excel, PowerPoint, or Word.":::

1. Press <kbd>F5</kbd>. Or select **Debug** > **Start Debugging**.
1. In the Office application, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.
1. Select **Get OneDrive File Names**. If you're logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane. If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to sign in. After you sign in, the file and folder names appear.

## Deploy the add-in

When you're ready to deploy to a staging or production server, be sure to update the following areas in the project solution.

- In the **appsettings.json** file, change the **domain** to your staging or production domain name.
- Update any references to `localhost:7080` throughout your project to use your staging or production URL.
- Update any references to `localhost:7080` in your Azure App registration, or create a new registration for use in staging or production.

For more information, see [Host and deploy ASP.NET Core](/aspnet/core/host-and-deploy/).

## See also

- [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).
- [Authorize to Microsoft Graph with SSO](authorize-to-microsoft-graph.md).
