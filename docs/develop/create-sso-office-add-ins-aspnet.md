# Create an ASP.NET Office Add-in that uses Single Sign-on (preview)

Users can sign into Office and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign-on a second time. For an overview, see [Single Sign-on to Office, your Office Web Add-in, and Microsoft Graph (preview)](..\docs\develop\sso-in-office-add-ins.md) .

This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET. 

> Note: For a similar article about a NodeJS-based add-in, see [Create a NodeJS Office Add-in that uses Single Sign-on](..\docs\develop\create-sso-office-add-ins-nodejs.md) .

## Prerequisites

* Visual Studio 2017 Version 15.3 (26424.2-Preview) or later.

* Office 2016, Version 1704,  build 8027.nnnn or later. (The Office 365 subscription version, sometimes called “Click to Run”.)  You many need to be an Office Insider to obtain this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1) .

## Setup the starter project

1. Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso). 

1. Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. 

    > Note: There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the *.sln file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side..

1. After the project opens, build it in Visual Studio, which will cause Visual Studio to install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.

1. After the project has completely built, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon. 

1. Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane. The button in the task pane is not wired up yet. 
2. In Visual Studio, stop the debugger.

## Register the add-in with Azure AD V2 endpoint

1. Navigate to [https://apps.dev.microsoft.com/?test=build2017](https://apps.dev.microsoft.com/?test=build2017) . 

1. Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com

1. Click **Add an app**.

1. When prompted, use “Office-Add-in-ASPNET-SSO” as the app name, and then press **Create application**.

1. When the configuration page for the app opens, copy the **Application Id** and save it. You will use it in a later procedure. 

    > Note: This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You will need it in a later procedure. Then close the dialog.

1. In the **Platforms** section, click **Add Platform**. 

1. In the dialog that opens, select **Web API**.

1. An **Application ID URI** has been generated of the form “api://{App ID GUID}”. Replace the GUID with “localhost:44355”. The entire ID should read `api://localhost:44355`. (The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match. It should read `api://localhost:44355/access_as_user`.)

1. In the **Pre-authorized applications** section, there is an empty **Application ID** box. Enter the following ID in the box (this is the ID of Microsoft Office):  `d3590ed6-52b3-4102-aeff-aad2292ab01c`.

1. Open the **Scope** drop down beside the **Application ID** and check the box for `api://localhost:44355/access_as_user`.

1. Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.

1. In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://localhost:44355`. 

    > Note: As of this writing, the **Web API** platform sometimes disappears from the **Platforms** section, particularly if the page is refreshed after the **Web** platform is added *and the registration page is saved*. For reassurance that your **Web API** platform is still part of the registration, click the **Edit Application Manifest** button near the bottom of the page. You should see the `api://localhost:44355` string in the **identifierUris** property of the manifest. There will also be a **oauth2Permissions** property whose **value** subproperty has the value `access_as_user`.

1. Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.

1. In the dialog, check the boxes for the following permissions (some may already be checked by default): 
 * Files.Read.All
 * offline_access
 * openid
 * profile

1. Click **OK** at the bottom of the dialog.

1. Click **Save** at the bottom of the registration page.

## Grant admin consent to the add-in

1. If the add-in is not running in Visual Studio, press F5 to run it. It needs to be running in IIS for this procedure to complete smoothly. 

1. In the following string, replace the placeholder “{application_ID}” with the Application ID that you copied when you registered your add-in.

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Paste the resulting URL into a browser address bar and navigate to it.

1. When prompted, sign-in with the admin credentials to your Office 365 tenancy.

1. You are then prompted to grant permission for your add-in to access your Microsoft Graph data. Click **Accept**. 

1. The browser window/tab is then redirected to the **Redirect URL** that you specified when you registered the add-in, so the home page of the add-in opens in the browser. 

2. In the browser's address bar you'll see a "tenant" query parameter with a GUID value. This is the ID of your Office 365 tenancy. Copy and save this value. You will use it in a later step.

3. Close the window/tab.

1. Stop the debugger in Visual Studio.

## Configure the add-in

1. In the following string, replace the placeholder “{tenant_ID}” with the Office 365 tenant ID you obtained earlier. If for any reason, you didn't get the ID earlier, use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.

1. Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.

1. Give the following values to the corresponding keys:

|Key|Value|
|:-----|:-----|
|ida:ClientID|The application ID you obtained when you registered the add-in.|
|ida:Audience|The application ID you obtained when you registered the add-in.|
|ida:Password|TThe password you obtained when you registered the add-in.|


Here’s an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*.

```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
```

> Note: Leave the other settings in the **appSettings** section unchanged.


1. Save and close the file.

1. In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.

1. Scroll to the bottom of the file.

1. Just above the end </VersionOverrides> tag, you will find the following markup:

    ```xml
    <WebApplicationId>{application_GUID here}</WebApplicationId>
    <WebApplicationResource>api://localhost:44355<WebApplicationResource>
    <WebApplicationScopes>
        <WebApplicationScope>profile</WebApplicationScope>
        <WebApplicationScope>openid</WebApplicationScope>
        <WebApplicationScope>offline_access</WebApplicationScope>
        <WebApplicationScope>files.read.all</WebApplicationScope>
    </WebApplicationScopes>
    ```

1. Replace the placeholder “{application_GUID here}” in the markup with the Application ID that you copied when you registered your add-in. This is the same ID you used in for the ClientID and Audience in the web.config.

    >Note: 
    >
    >* The **WebApplicationResource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.
    >* The **WebApplicationScopes** section is used only to generate a consent dialog if the add-in is sold through the Office Store.

1. Save and close the file.

## Code the client side

1. Open the Home.js file in the **Scripts** folder. It already has some code in it:

    * An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.
    * A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.

1. Below the assignment to `Office.initialize`, add the code below. Note the following about this code: 

 * The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token. 
 * If no user is signed into Office, the Office host will prompt the user to sign in. 
 * The options parameter sets `forceConsent` to false, so the user will not be prompted to consent to giving the Office host access to your add-in.

```js
function getOneDriveItems() {
Office.context.auth.getAccessTokenAsync({ forceConsent: false },
    function (result) {
        if (result.status === "succeeded") {
            // TODO1: Use the access token to get Microsoft Graph data.
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
        console.log(result.error);
    });
}
```

1. Save and close the file.

## Code the server side

### Configure the OWIN middleware

1. Open the Startup.cs file in the root of the project. 

1. Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:

    `public partial class Startup`

1. Add the following line to the body of the `Configure` method. You create the `ConfigureAuth` method in a later step.

    `ConfigureAuth(app);`

1. Save and close the file.

1. Right-click the **App_Start** folder and select **Add | Class**. 

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
	// TODO2: Configure the validation settings
	// TODO3: Specify the type of authorization and the discovery endpoint
	// of the secure token service.
}
```

1. Replace the TODO2 with the following. Note:

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

1. Replace TODO3 with the following. Note:

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

1. Ensure that the following `using` statements are at the top of the file.

```
using Microsoft.Identity.Client;
using System.IdentityModel.Tokens;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;
```

1. Just above the line that declares the `ValuesController`, add the attribute `[Authorize]`. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called; so only callers with a valid access token to your add-in can invoke the methods of the controller. 

1. Add the following method to the `ValuesController`:

```
// GET api/values
public async Task<IEnumerable<string>> Get()
{
    // TODO4: Validate the scopes of the access token.
}
```

1. Replace TODO4 with the following code to validate that the scopes that are specified in the token include `access_as_user`. 

```
string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
if (addinScopes.Contains("access_as_user"))
{
    // TODO5: Get the raw token that the add-in page received from the Office host.
    // TODO6: Get the access token for MS Graph.
    // TODO7: Get the names of files and folders in OneDrive for Business by using the Microsoft Graph API.
    // TODO8: Remove excess information from the data and send the data to the client.
}
return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
```

1. Replace TODO5 with the following code which turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.

```var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
```

1. Replace TODO6 with the following code. Note:

 * Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object. 
 * The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.
 * The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.

```
ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
string[] graphScopes = { "profile", "Files.Read.All" };
AuthenticationResult result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
```

1. Replace TODO7 with the following. Note:

 * The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.
 * Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first 3 folder of file names.

```
var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
```

1. Replace TODO8 with the following. Note that although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.

```List<string> itemNames = new List<string>();
      foreach (OneDriveItem item in getFilesResult)
      {
          itemNames.Add(item.Name);
      }                    
      return itemNames;
```

## Run the add-in

1. Make sure you have some files in your OneDrive for Business.

1. In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon. 

1. Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane. 

1. Press the button **Get My Files from** OneDrive. If you are not signed into Office, you will be prompted to sign in.

1. After you are signed in, a list of your files and folders on OneDrive for Business will appear below the button. This may take over 15 seconds, especially the first time. 



