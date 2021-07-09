---
title: Get the ID token of the signed-in user
description: Call the getAccessToken API to get the ID token with name, email, and additional information about the signed in user.
ms.date: 07/15/2021
localization_priority: Normal
---

# Get the ID token of the signed-in user

Use the `getAccessToken` API to get an ID token for the user that is signed in to Office. The user ID token contains information about the signed in user, such as their name and email. You can also obtain a unique ID from the ID token to identify the user when calling your own web services. To call `getAccessToken` you must configure your Office Add-in to use SSO with Office.

In this article you'll create an Office Add-in that gets the ID token, and displays the user's name, email, and unique ID in the task pane.

> [!NOTE]
> SSO with Office and the `getAccessToken` API do not work in all scenarios. You should always implement a fallback dialog to sign in the user when SSO is unavailable. For more information see TBD.

## Create an app registration

To use SSO with Office, you need to create an app registration in the Azure portal so the Microsoft identity platform can provide authentication and authorization services for your Office Add-in and its users.

[!INCLUDE [create-azure-app-registration-for-sso](../includes/create-azure-app-registration-for-sso.md)]

## Create the Office Add-in

# [Visual Studio 2019](#tab/vs2019)

1. Start Visual Studio 2019 and choose to **Create a new project**.
1. Search for and select the **Excel Web Add-in** project template. Then choose **Next**. Note: SSO works with any Office application, but for this article will work with Excel.
1. Enter a project name, such as **sso-display-user-info** and choose **Create**. You can leave the other fields at default values.
1. In the **Choose the add-in type** dialog box, select **Add new functionality to Excel**, and choose **Finish**.

The project is created and will contain two projects in the solution.
- **sso-display-user-info**: Contains the manifest and details for sideloading the add-in to Excel.
- **sso-display-user-infoWeb**: The ASP.NET project that hosts the web pages for the add-in.

# [yo office](#tab/yooffice)

Be sure you have [Set up your development environment](../overview/set-up-your-dev-environment.md).

1. Enter the following command to create the project.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

The project is created in a new folder named **sso-display-user-info**.
---
## Configure the manifest

1. In **Solution Explorer** open **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**
1. Near the bottom of the manifest is a closing `</Resources>` element. Insert the following XML just below the `</Resources>` element but before the closing `</VersionOverrides>` element.

   ```xml
   <WebApplicationInfo>
       <Id>[ApplicationID]</Id>
       <Resource>api://localhost:[port]/[ApplicationID]</Resource>
       <Scopes>
           <Scope>openid</Scope>
           <Scope>user.read</Scope>
           <Scope>profile</Scope>
       </Scopes>
   </WebApplicationInfo>
   ```

1. Replace **[port]** with the correct port number for your project.
1. Replace both **[Application ID]** placeholders with the actual application ID from your app registration.
1. Save the file.

The XML you inserted contains the following elements and information.

- **WebApplicationInfo** - The parent of the following elements.
- **Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).
- **Resource** - The URL of the add-in. This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD. The domain part of this URI must match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest and the URI must end with the client ID in the `<Id>`.
- **Scopes** - The parent of one or more **Scope** elements.
- **Scope** - Specifies a permission that the add-in needs to AAD. The `profile` and `openID` permissions are always needed and may be the only permissions needed, if your add-in does not access Microsoft Graph. If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).

For Office applications other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.

## Add the jwt-decode package

You can call the `getAccessToken` API to get the ID token from Office. First lets add the jwt-decode package to make it easier to decode and view the ID token.

# [Visual Studio 2019](#tab/vs2019)

1. Open the Visual Studio solution.
1. On the menu, choose **Tools > NuGet Package Manager > Package Manager Console**.
1. Enter the following command in the **Package Manager Console**.
    
    `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# [yo office](#tab/yooffice)

1. From a terminal/console window go to the root folder for your add-in project.
1. Enter the following command
    
    `npm install jwt-decode`
---
## Add UI to the task pane

We need to modify the task pane so that it can display the user information we'll get from the ID token.

# [Visual Studio 2019](#tab/vs2019)

1. Open the Home.html file.
1. Add the following script tag to the `<head>` section of the page. This will include the jwt-decode package we added earlier.
    
    ```html
    <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
    ```
    
1. Replace the body with the following HTML.
    
    ```html
    <body>
        <h1>Welcome</h1>
        <p>Sign in to Office, then choose the <b>Get ID Token</b> button to see your ID token information.</p>
        <button id="getIDToken">Get ID Token</button>
        <div>
            <span id='userInfo'></span>
        </div>
        </main>
    </body>
    ```
    

# [yo office](#tab/yooffice)

1. Open the taskpane.html file.
1. Replace the body with the following HTML.
    
    ```html
    <body>
        <h1>Welcome</h1>
        <p>Sign in to Office, then choose the <b>Get ID Token</b> button to see your ID token information.</p>
        <button id="getIDToken">Get ID Token</button>
        <div>
            <span id='userInfo'></span>
        </div>
        </main>
    </body>
    ```
---
## Call the getAccessToken API

The final step is to get the ID token by calling `getAccessToken`.

# [Visual Studio 2019](#tab/vs2019)

1. Open the Home.js file.
1. Replace the entire contents of the file with the following code.
    
    ```javascript
    (function () {
        "use strict";
    
        // The initialize function must be run each time a new page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#getIDToken').click(getIDToken);
            });
        };
    
        async function getIDToken() {
            try {
                let userTokenEncoded = await OfficeRuntime.auth.getAccessToken();
                let userToken = jwt_decode(userTokenEncoded);
                document.getElementById('userInfo').innerHTML = "name: " + userToken.name + "<br> email: " + userToken.preferred_username + "<br> id: " + userToken.oid;
                console.log(userToken);
            } catch (error) {
                console.error(error);
            }
        }
    
    })();
    ```
    
1. Save the file.

# [yo office](#tab/yooffice)

1. Open the taskpane.js file.
1. Replace the entire contents of the file with the following code.
    
    ```javascript
    import jwt_decode from 'jwt-decode';
    
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        document.getElementById("getIDToken").onclick = getIDToken;
      }
    });
    
    async function getIDToken() {
      try {
        let userTokenEncoded = await OfficeRuntime.auth.getAccessToken();
        let userToken = jwt_decode(userTokenEncoded);
        document.getElementById('userInfo').innerHTML = "name: " + userToken.name + "<br> email: " + userToken.preferred_username + "<br> id: " + userToken.oid;
        console.log(userToken);
      } catch (error) {
        console.error(error);
      }
    }
    ``` 
    
1. Save the file.
---

run it!


## Next steps


