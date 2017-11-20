---
title: Create an ASP.NET Office Add-in that uses single sign-on
description: 
ms.date: 11/20/2017 
---

# Create an ASP.NET Office Add-in that uses single sign-on

When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.

> [!NOTE]
> For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).

## Prerequisites

* The latest available version of Visual Studio 2017 Preview.

  > [!NOTE]
  > The latest version of Visual Studio 2017 Preview is not currently compatible with the add-in manifest markup that is required for SSO. Details about how to work around this are provided in the following procedures.

* Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”). You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## Set up the starter project

1. Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).

1. Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.

    > [!NOTE]
    > There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the *.sln file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side..

1. After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.

    > [!IMPORTANT]
    > The packages.config in the root of the web API project, specifies version `1.1.1-alpha0393` of Microsoft.Identity.Client, the MSAL library. You should verify that this version (or later) gets installed after you press F5 for the first time: On the **Tools** menu, navigate to **Nuget Package Manager** > **Manage Nuget Packages for Solution** > **Installed**. Scroll to **Microsoft.Identity.Client** to see the installed version. If it is earlier than `1.1.1-alpha0393` (or does not appear on the **Installed** list), then navigate to **Nuget Package Manager** > **Package Manager Console**. At the console, run the command `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`.

1. After the project has completely built, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.

1. Press the **Show add-in** button in this group to see the add-in’s UI in the task pane. The button in the task pane is not wired up yet.
2. In Visual Studio, stop the debugger.

## Register the add-in with Azure AD v2.0 endpoint

1. Navigate to [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com

1. Click **Add an app**.

1. When prompted, use “Office-Add-in-ASPNET-SSO” as the app name, and then press **Create application**.

1. When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.

    > [!NOTE]
    > This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.

1. In the **Platforms** section, click **Add Platform**.

1. In the dialog that opens, select **Web API**.

1. An **Application ID URI** has been generated of the form “api://{App ID GUID}”. Insert the string “localhost:44355/” between the double forward slashes and the GUID. The entire ID should read `api://localhost:44355/{App ID GUID}`. (The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match. It should read `api://localhost:44355/{App ID GUID}/access_as_user`.)

1. In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized. Each time you enter one, a new empty textbox appears. (Enter only the GUID.)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. Open the **Scope** drop-down beside each **Application ID** and check the box for `api://localhost:44355/{App ID GUID}/access_as_user`.

1. Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.

1. In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://localhost:44355`.

    > [!NOTE]
    > As of this writing, the **Web API** platform sometimes disappears from the **Platforms** section, particularly if the page is refreshed after the **Web** platform is added *and the registration page is saved*. For reassurance that your **Web API** platform is still part of the registration, click the **Edit Application Manifest** button near the bottom of the page. You should see the `api://localhost:44355/{App ID GUID}` string in the **identifierUris** property of the manifest. There will also be a **oauth2Permissions** property whose **value** subproperty has the value `access_as_user`.

1. Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.

1. In the dialog box, check the boxes for the following permissions (some may already be checked by default). Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`. The `profile` permission is required for the Office host to get a token to your add-in web application.
    * Files.Read.All
    * offline_access
    * openid
    * profile

1. At the bottom of the dialog, click **OK**.

1. At the bottom of the registration page, click **Save**.

## Grant admin consent to the add-in

> [!NOTE]
> This procedure is only needed when you're developing the add-in. When your production add-in is deployed to the Office Store or an add-in catalog, users will individually trust it or an admin will consent for organization at installation.

1. If the add-in isn't running in Visual Studio, press **F5** to run it. It needs to be running in IIS for this procedure to complete smoothly.

1. In the following string, replace the placeholder “{application_ID}” with the Application ID that you copied when you registered your add-in:
    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Paste the resulting URL into a browser address bar and navigate to it.

1. When prompted, sign in with the admin credentials to your Office 365 tenancy.

1. You are then prompted to grant permission for your add-in to access your Microsoft Graph data. Click **Accept**.

1. The browser window/tab is then redirected to the **Redirect URL** that you specified when you registered the add-in, so the home page of the add-in opens in the browser.

2. In the browser's address bar you'll see a "tenant" query parameter with a GUID value. This is the ID of your Office 365 tenancy. Copy and save this value. You'll use it in a later step.

3. Close the window/tab.

1. Stop the debugger in Visual Studio.

## Configure the add-in

1. In the following string, replace the placeholder “{tenant_ID}” with the Office 365 tenant ID you obtained earlier. If for any reason, you didn't get the ID earlier, use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.

1. Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.

1. Assign the following values to the corresponding keys:

    |Key|Value|
    |:-----|:-----|
    |ida:ClientID|The application ID you obtained when you registered the add-in.|
    |ida:Audience|The application ID you obtained when you registered the add-in.|
    |ida:Password|TThe password you obtained when you registered the add-in.|


   The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup will be more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource - relative to the Office host - and an OAuth client - relative to Microsoft Graph.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    ```

   > [!NOTE]
   > Leave the other settings in the **appSettings** section unchanged.

1. Save and close the file.

1. In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.

1. Scroll to the bottom of the file.

1. Just above the end `</VersionOverrides>` tag, you'll find the following markup:

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}<Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in. The "{}" are not part of the ID, so do not include them. This is the same ID you used in for the ClientID and Audience in the web.config.

    > [!NOTE]
    > * The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.
    > * The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through the Office Store.

1. Open the **Warnings** tab of the **Error List** in Visual Studio. If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not  recognize the SSO markup. As a workaround, do the following for a Word, Excel, or PowerPoint add-in. (If you are working with an Outlook add-in see the workaround below.)

   - **Workaround for Word, Excel, and Powerpoint**

        1. Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.

        2. Press F5 to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.

        4. Save the copy of the manifest.

        5. Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).

        6. Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.

        7. Expand **Configuration Properties** and select **Configuration**.

        8. Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).

        9. Press **OK** to close the dialog box.

   - **Workaround for Outlook**

        1. On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.

        2. Rename the existing file to `MailAppVersionOverridesV1_1.old`.

        3. Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Save and close the main manifest file in Visual Studio.

## Code the client side

1. Open the Home.js file in the **Scripts** folder. It already has some code in it:
    * An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.
    * A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.

1. Below the assignment to `Office.initialize`, add the code below. Note the following about this code:

    * The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.
    * If no user is signed into Office, the Office host will prompt the user to sign in.
    * The options parameter sets `forceConsent` to false, so the user will not be prompted to consent to giving the Office host access to your add-in.

    ```js
    function getOneDriveFiles() {
        getDataWithToken({ forceConsent: false });
    }

    function getDataWithToken(options) {
        Office.context.auth.getAccessTokenAsync(options,
            function (result) {
                if (result.status === "succeeded") {
                    TODO1: Use the access token to get Microsoft Graph data.
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }
    ```

1. Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Below the `getOneDriveFiles` method, add the following. This utility method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            TODO2: Handle errors and the case where Microsoft Graph
                   requires additional form of authentication.
        });
    }
    ```

1. Replace the TODO2 with the following lines. Note the following about this code:

    * When the failure is because Microsoft Graph requires an additional form of authentication, the `exceptionMessage` will be a JSON string containing "capolids". In that case, the Office host needs to get a new token.  
    * The exception message tells AAD to prompt the user for all required forms of authentication, so it must be passed to the Office host, which in turn passes it to AAD when it asks for a new token.
    * The `authChallenge` option is the method of passing this string to the Office host.
    * If the error is something other than a request for additional authentication, it is logged to the console.

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    if (exceptionMessage.indexOf("capolids") !== -1) {
        getDataWithToken({ authChallenge: exceptionMessage });
    } else {
        console.log(result.error);
    }
    ```

1. Save and close the file.

## Code the server side

### Configure the OWIN middleware

1. Open the Startup.cs file in the root of the project.

1. Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:

    `public partial class Startup`

1. Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.

    `ConfigureAuth(app);`

1. Save and close the file.

1. Right-click the **App_Start** folder and select **Add > Class**.

1. In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.

1. Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Ensure that all of the following `using` statements are at the top of the file.

    ```
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:

    `public partial class Startup`

1. Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.

    ```
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Replace the TODO3 with the following. Note:

    * The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.
    * Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.
    * Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.

    ```
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Replace TODO4 with the following. Note:

    * The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.
    * The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.

    ```
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
            });
    ```

1. Save and close the file.

### Create the /api/values controller

1. Open the file **Controllers\ValueController.cs**.

2. Ensure that the following `using` statements are at the top of the file.

    ```
    using Microsoft.Identity.Client;
    using System;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web;
    using System.Web.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.

4. Add the following method to the `ValuesController`:

    ```
    // GET api/values
    public async Task<IEnumerable<string>> Get()
    {
        // TODO5: Validate the scopes of the access token.
    }
    ```

5. Replace TODO5 with the following code to validate that the scopes that are specified in the token include `access_as_user`.

    ```
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO6: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO7: Get the access token for Microsoft Graph.
        // TODO8: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO9: Remove excess information from the data and send the data to the client.
    }
    return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
    ```

    > [!NOTE]
    > You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.

6. Replace TODO6 with the following code. Note:
    * It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.
    * Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.
    * The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.
    * MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.

    ```
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. Replace TODO7 with the following code. Note:

    * The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.
    * If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.
    * The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.
    * Any exceptions that are not of type `MsalUiRequiredException` are intentionally not caught, so they will propagate to the client.

    ```
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalUiRequiredException e)
    {        
        if (String.IsNullOrEmpty(e.Claims))
        {
            throw e;
        }
        else
        {
            throw new HttpException(e.Claims);
        }   
    }
    ```

8. Replace TODO8 with the following. Note:

    * The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.
    * Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first 3 folder or file names.

    ```
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    ```

9. Replace TODO9 with the following. Note that although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.

    ```
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in getFilesResult)
    {
      itemNames.Add(item.Name);
    }                    
    return itemNames;
    ```

## Run the add-in

1. Ensure that you have some files in your OneDrive so that you can verify the results.

1. In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.

1. Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.

1. Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.
    
    > [!NOTE]
    > If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.

1. After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.
