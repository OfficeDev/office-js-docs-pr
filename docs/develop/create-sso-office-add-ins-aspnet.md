---
title: Create an ASP.NET Office Add-in that uses single sign-on
description: 
ms.date: 01/23/2018
localization_priority: Priority
---

# Create an ASP.NET Office Add-in that uses single sign-on (preview)

When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.

> [!NOTE]
> For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).

## Prerequisites

* The latest available version of Visual Studio 2017.

* Office 365 (subscription version, also called “Click to Run”). Latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Note: when a version/build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that version/build.

## Set up the starter project

1. Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).

1. Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.

    > [!NOTE]
    > There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.

1. After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.

    > [!NOTE]
    > You will get an error about the Identity namespace. This is a side effect of a configuration issue that you will fix with the next step. The important thing is that the packages are installed.

1. Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.4-preview0002`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.

   > 1. On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**. 

   > 2. At the console, run the following command. It may take a minute or more to complete even with a fast Internet connection. When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.4-preview0002' ...** near the end of the output in the console.

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`

   > 3. In **Solution Explorer**, expand **References** of **Office-Add-in-ASPNET-SSO-WebAPI** project. Verify that **Microsoft.Identity.Client** is listed. If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference Wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**

1. Build the project a second time.

## Register the add-in with Azure AD v2.0 endpoint

The following instruction are written generically so they can be used in multiple places. For this ariticle do the following:
- Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.
- Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.
- When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions. Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`. The `profile` permission is required for the Office host to get a token to your add-in web application.
    * Files.Read.All
    * offline_access
    * openid
    * profile


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## Grant administrator consent to the add-in

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## Configure the add-in

1. In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenant ID. Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.

3. Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.

4. Assign the following values to the corresponding keys:

    |Key|Value|
    |:-----|:-----|
    |ida:ClientID|The application ID you obtained when you registered the add-in.|
    |ida:Audience|The application ID you obtained when you registered the add-in.|
    |ida:Password|The password you obtained when you registered the add-in.|

   The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.

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
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
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
    > * The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.

1. Open the **Warnings** tab of the **Error List** in Visual Studio. If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not recognize the SSO markup. As a workaround, do the following for a Word, Excel, or PowerPoint add-in. (If you are working with an Outlook add-in see the workaround below.)

   - **Workaround for Word, Excel, and Powerpoint**

        1. Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.

        2. Press **F5** to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

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
    * A `logErrors` method that will log to console errors that are not intended for the end user.

1. Below the assignment to `Office.initialize`, add the code below. Note the following about this code:

    * The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token. 
    * You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }	
    ```

1. Below the `getOneDriveFiles` method, add the code below. Note the following about this code:

    * The [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.
    * If no user is signed into Office, the Office host will prompt the user to sign in.
    * The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in. The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.
    * You will create the `handleClientSideErrors` method in a later step.

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Below the `getOneDriveFiles` method, add the following. About this code, note:

    * This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.
    * You will create the `handleServerSideErrors` method in a later step.

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### Create the error-handling methods

1. Below the `getData` method, add the following method. This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service. These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. Replace `TODO2` with the following code. Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor. In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Replace `TODO3` with the following code. Error 13002 occurs when user's sign-in or consent was aborted. Ask the user to try again but no more than once again.

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. Replace `TODO4` with the following code. Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft account. Ask the user to sign-out and then in again with a supported account type.

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > Errors 13004 and 13005 are not handled in this method because they should only occur in development. They cannot be fixed by runtime code and there would be no point in reporting them to an end user.

1. Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. Replace `TODO6` with the following code. Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application. This may be a temporary network issue. Ask the user to try again later.

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. Replace `TODO8` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. Replace `TODO9` with the following code.

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle missing consent and scope (permission) related issues.

        // TODO13: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO14: Log all other server errors.
    }
    ```

1. Replace `TODO10` with the following code. Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data. However, when AAD sends a message to the add-in's web service asking for an additonal authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed. The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object. Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object. This value will be in the **Message** property, so the code needs to parse out that property as well.

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. Replace `TODO11` with the following code. Note about this code:

    * Error 50076 occurs when Microsoft Graph requires an additional form of authentication.
    * The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication. 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. Replace `TODO12` with the following code. You will replace the three `TODO`s in this code with an *inner* conditional block in the next few steps.

    ```javascript
    else if (exceptionMessage) {

        // TODO12A: Handle the case where consent has not been granted, or has been revoked.

        // TODO12B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO12C: Handle the case where the token that the add-in's client-side sends to it's 
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. Replace `TODO12A` with the following code. (This creates the first part of an *inner* conditional block.) Note about this code:

    * Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions. 
    * The add-in should get a new token with the `forceConsent` option set to `true`.

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. Replace `TODO12B` with the following code. Note about this code:

    * Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.
    * The add-in should report the error.

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. Replace `TODO12C` with the following code. Note about this code:

    * Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.
    * The add-in should report the error.

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. Replace `TODO13` with the following code. (This is part of the *outer* conditional block and should be immediately after the close bracket of the structure that begins with `else if (exceptionMessage) {` and at the same level of indentation.) Note about this code:

    * The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`. Server-side code you will create in a latter step will relay this message to the add-in's client.
    * In this case, the add-in should start the entire authentication process over by resetting the counter and flag varibles, and then re-calling the button handler method.

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. Replace `TODO14` with the following code.

    ```javascript
    else {
        logError(result);
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

    ```csharp
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

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Replace the TODO3 with the following. Note about this code:

    * The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.
    * Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.
    * Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Replace TODO4 with the following. Note about this code:

    * The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.
    * The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. Save and close the file.

### Create the /api/values controller

1. Open the file **Controllers\ValueController.cs**.

2. Ensure that the following `using` statements are at the top of the file.

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.

    > [!NOTE]
    > A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom **FilterAttribute** classes. This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed. This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).    

4. Add the following method to the `ValuesController`. Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method. This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client. 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.

6. Replace `TODO2` with the following code. Note about this code:
    * It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.
    * Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.
    * The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.
    * MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. Replace `TODO3` with the following code. Note about this code:

    * The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.
    * If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.
    * The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.
    * Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. Replace `TODO3a` with the following code. Note about this code:

    * If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information. 
    * The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.
    * The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.
    * The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.
    * You will create the `SendErrorToClient` method in a later step. It's second parameter is an **Exception** object. In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. Replace `TODO3b` and `TODO3c` with the following code. Note about this code:

    * If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked). AAD will return "400 Bad Request" with error `AADSTS65001`. MSAL throws a **MsalUiRequiredException** with this information. The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.
    *  If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`. MSAL throws a **MsalUiRequiredException** with this information. The client should inform the user.
    *  The entire description is included beause 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope. 
    *  The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.
    *  There is no custom message, so `null` is passed for the third parameter.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. Replace `TODO3d` with the following code. Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401). The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. Replace `TODO4` with the following. Note about this code:

    * The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.
    * Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.
    * If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken". ASP.NET then throws a **RuntimeBinderException**. This is also what happens when the token is expired, although MSAL should prevent that from ever happening. 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. Replace `TODO5` with the following. Note about this code: 

    * Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.
    * The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. Below the Get method, add the following method. About this code note:  

    * The method relays to the client information about a server-side exception. 
    * If the original exception is passed to the method, then the HttpError constuctor will include information from the exception object in an **ExceptionMessage** property.  
    * If `null` is passed for the exception, then the HttpError constuctor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## Run the add-in

1. Ensure that you have some files in your OneDrive so that you can verify the results.

1. In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.

1. Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.

1. Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.
    
    > [!NOTE]
    > If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.

1. After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.
