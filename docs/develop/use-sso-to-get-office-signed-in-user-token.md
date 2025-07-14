---
title: Use SSO to get the identity of the signed-in user
description: Call the getAccessToken API to get the ID token with name, email, and additional information about the signed-in user.
ms.date: 06/24/2025
ms.localizationpriority: medium
---

# Use SSO to get the identity of the signed-in user

Use the `getAccessToken` API to get an access token that contains the identity for the current user signed in to Office. The access token is also an ID token because it contains identity claims about the signed-in user, such as their name and email. You can also use the ID token to identify the user when calling your own web services. To call `getAccessToken`, you must configure your Office Add-in to use SSO with Office.

In this article, you'll create an Office Add-in that gets the ID token, and displays the user's name, email, and unique ID in the task pane.

> [!NOTE]
> SSO with Office and the `getAccessToken` API don't work in all scenarios. Always implement a fallback dialog to sign in the user when SSO is unavailable. For more information, see [Authenticate and authorize with the Office dialog API](auth-with-office-dialog-api.md).

## Create an app registration

To use SSO with Office, you need to create an app registration in the Azure portal so the Microsoft identity platform can provide authentication and authorization services for your Office Add-in and its users.

1. To register your app, go to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page.

1. Sign in with the **_admin_** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Office-Add-in-SSO`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.
    - Set the application type to **Web** and then set **Redirect URI** to `https://localhost:[port]/dialog.html`. Replace `[port]` with the correct port number for your web application. If you created the add-in using Yo Office, the port number is typically 3000 and found in the package.json file. If you created the add-in with Visual Studio 2019, the port is found in the **SSL URL** property of the web project.
    - Choose **Register**.

1. On the **Office-Add-in-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**. You'll use both of them in later procedures.

    > [!NOTE]
    > This **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It's also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. Select **Authentication** under **Manage**. In the **Implicit grant** section, enable the checkboxes for both **Access token** and **ID token**.

1. Select **Save** at the top of the form.

1. Select **Expose an API** under **Manage**. Select the **Set** link. This will generate the Application ID URI in the form `api://[app-id-guid]`, where `[app-id-guid]` is the **Application (client) ID**.

1. In the generated ID, insert `localhost:[port]/` (note the forward slash "/" appended to the end) between the double forward slashes and the GUID. Replace `[port]` with the correct port number for your web application. If you created the add-in using Yo Office, the port number is typically 3000 and found in the package.json file. If you created the add-in with Visual Studio 2019, the port is found in the **SSL URL** property of the web project.

    When you're finished, the entire ID should have the form `api://localhost:[port]/[app-id-guid]`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **\<Scope\>** name.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user. Suggestions:

   - **Admin consent display name**: Office can act as the user.
   - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.
   - **User consent display name**: Office can act as you.
   - **User consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope** .

   > [!NOTE]
   > The domain part of the **\<Scope\>** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Authorized client applications** section, enter the following ID to pre-authorize all Microsoft Office application endpoints.

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (All Microsoft Office application endpoints)

    > [!NOTE]
    > The `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID pre-authorizes Office on all the following platforms. Alternatively, you can enter a proper subset of the following IDs if for any reason you want to deny authorization to Office on some platforms. Just leave out the IDs of the platforms from which you want to withhold authorization. Users of your add-in on those platforms will not be able to call your Web APIs, but other functionality in your add-in will still work.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. Select the **Add a client application** button and then, in the panel that opens, set the `[app-id-guid]` to the Application (client) ID and check the box for `api://localhost:44355/[app-id-guid]/access_as_user`.

1. Select **Add application**.

1. Select **API permissions** under **Manage** and select **Add a permission**. On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search for the permissions your add-in needs. Search for and select the **profile** permission. The `profile` permission is required for the Office application to get a token to your add-in web application.

   - profile

   > [!NOTE]
   > The `User.Read` permission may already be listed by default. It's a good practice not to ask for permissions that aren't needed, so we recommend that you uncheck the box for this permission if your add-in doesn't actually need it.

1. Select the **Add permissions** button at the bottom of the panel.

1. On the same page, choose the **Grant admin consent for \<tenant-name\>** button, and then select **Yes** for the confirmation that appears.

## Create the Office Add-in

# [Visual Studio 2019](#tab/vs2019)

1. Start Visual Studio 2019 and choose to **Create a new project**.
1. Search for and select the **Excel Web Add-in** project template. Then choose **Next**. Note: SSO works with any Office application, but Excel is the application being used with this article.
1. Enter a project name, such as **sso-display-user-info**, and choose **Create**. You can leave the other fields at default values.
1. In the **Choose the add-in type** dialog box, select **Add new functionality to Excel**, and choose **Finish**.

The project is created and will contain two projects in the solution.

- **sso-display-user-info**: Contains the manifest and details for sideloading the add-in to Excel.
- **sso-display-user-infoWeb**: The ASP.NET project that hosts the web pages for the add-in.

# [Yo Office](#tab/yooffice)

Be sure you have [Set up your development environment](../overview/set-up-your-dev-environment.md).

1. Enter the following command to create the project.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

The project is created in a new folder named **sso-display-user-info**.

---

## Configure the manifest

# [Visual Studio 2019](#tab/vs2019)

In **Solution Explorer**, open **sso-display-user-info** > **sso-display-user-infoManifest** > **sso-display-user-info.xml**.

# [Yo Office](#tab/yooffice)

In Visual Studio Code, open the **manifest.xml** file.

---

1. Near the bottom of the manifest is a closing `</Resources>` element. Insert the following XML just below the `</Resources>` element but before the closing `</VersionOverrides>` element. For Office applications other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.

    ```xml
    <WebApplicationInfo>
        <Id>[application-id]</Id>
        <Resource>api://localhost:[port]/[application-id]</Resource>
        <Scopes>
            <Scope>openid</Scope>
            <Scope>user.read</Scope>
            <Scope>profile</Scope>
        </Scopes>
    </WebApplicationInfo>
   ```

1. Replace `[port]` with the correct port number for your project. If you created the add-in using Yo Office, the port number is typically 3000 and found in the package.json file. If you created the add-in with Visual Studio 2019, the port is found in the **SSL URL** property of the web project.
1. Replace both `[application-id]` placeholders with the actual application ID from your app registration.
1. Save the file.

The XML you inserted contains the following elements and information.

- **\<WebApplicationInfo\>** - The parent of the following elements.
- **\<Id\>** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](register-sso-add-in-aad-v2.md).
- **\<Resource\>** - The URL of the add-in. This is the same URI (including the `api:` protocol) that you used when registering the add-in in Microsoft Entra ID. The domain part of this URI must match the domain, including any subdomains, used in the URLs in the **\<Resources\>** section of the add-in's manifest and the URI must end with the client ID in the **\<Id\>**.
- **\<Scopes\>** - The parent of one or more **\<Scope\>** elements.
- **\<Scope\>** - Specifies a permission that the add-in needs. The `profile` and `openID` permissions are always needed and may be the only permissions needed, if your add-in doesn't access Microsoft Graph. If it does, you also need **\<Scope\>** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires the `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).

## Add the jwt-decode package

You can call the `getAccessToken` API to get the ID token from Office. First, let's add the jwt-decode package to make it easier to decode and view the ID token.

# [Visual Studio 2019](#tab/vs2019)

1. Open the Visual Studio solution.
1. On the menu, choose **Tools** > **NuGet Package Manager** > **Package Manager Console**.
1. Enter the following command in the **Package Manager Console**.

    `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# [Yo Office](#tab/yooffice)

1. From a terminal/console window go to the root folder for your add-in project.
1. Enter the following command

    `npm install jwt-decode`

---

## Add UI to the task pane

Modify the task pane so that it can display the user information you'll get from the ID token.

# [Visual Studio 2019](#tab/vs2019)

1. Open the Home.html file.
1. Add the following script tag to the `<head>` section of the page. This will include the jwt-decode package was added earlier.

    ```html
    <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
    ```

1. Replace the `<body>` section with the following HTML.

    ```html
    <body>
      <h1>Welcome</h1>
      <p>
        Sign in to Office, then choose the <b>Get ID Token</b> button to see your
        ID token information.
      </p>
      <button id="getIDToken">Get ID Token</button>
      <div>
        <span id="userInfo"></span>
      </div>
    </body>
    ```

# [Yo Office](#tab/yooffice)

1. Open the **src/taskpane/taskpane.html** file.
1. Replace the `<body>` section with the following HTML.

    ```html
    <body>
      <h1>Welcome</h1>
      <p>
        Sign in to Office, then choose the <b>Get ID Token</b> button to see your
        ID token information.
      </p>
      <button id="getIDToken">Get ID Token</button>
      <div>
        <span id="userInfo"></span>
      </div>
    </body>
    ```

---

## Call the getAccessToken API

The final step is to get the ID token by calling `getAccessToken`.

# [Visual Studio 2019](#tab/vs2019)

1. Open the **Home.js** file.
1. Replace the entire contents of the file with the following code.

    ```javascript
    (function () {
      "use strict";

      // The initialize function must be run each time a new page is loaded.
      Office.initialize = function (reason) {
        $(document).ready(function () {
          $("#getIDToken").on("click", getIDToken);
        });
      };

      async function getIDToken() {
        try {
          let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
          });
          let userToken = jwt_decode(userTokenEncoded);
          document.getElementById("userInfo").innerHTML =
            "name: " +
            userToken.name +
            "<br>email: " +
            userToken.preferred_username +
            "<br>id: " +
            userToken.oid;
          console.log(userToken);
        } catch (error) {
          document.getElementById("userInfo").innerHTML =
            "An error occurred. <br>Name: " +
            error.name +
            "<br>Code: " +
            error.code +
            "<br>Message: " +
            error.message;
          console.log(error);
        }
      }
    })();
    ```

1. Save the file.

# [Yo Office](#tab/yooffice)

1. Open the **src/taskpane/taskpane.js** file.
1. Replace the entire contents of the file with the following code.

    ```javascript
    import jwt_decode from "jwt-decode";

    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        document.getElementById("getIDToken").onclick = getIDToken;
      }
    });

    async function getIDToken() {
      try {
        let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
        });
        let userToken = jwt_decode(userTokenEncoded);
        document.getElementById("userInfo").innerHTML =
          "name: " +
          userToken.name +
          "<br>email: " +
          userToken.preferred_username +
          "<br>id: " +
          userToken.oid;
        console.log(userToken);
      } catch (error) {
        document.getElementById("userInfo").innerHTML =
          "An error occurred. <br>Name: " +
          error.name +
          "<br>Code: " +
          error.code +
          "<br>Message: " +
          error.message;
        console.log(error);
      }
    }
    ```

1. Save the file.

---

## Run the add-in

# [Visual Studio 2019](#tab/vs2019)

Choose **Debug** > **Start Debugging**, or press <kbd>F5</kbd>.

# [Yo Office](#tab/yooffice)

Run `npm start` from the command line.

---

1. When Excel starts, sign in to Office with the same tenant account you used to create the app registration.
1. On the **Home** ribbon, choose **Show Taskpane** to open the add-in.
1. In the add-in's task pane, choose **Get ID token**.

The add-in will display the name, email, and ID of the account you signed in with.

> [!NOTE]
> If you encounter any errors, review the registration steps in this article for the app registration. Missing a detail when setting up the app registration is a common cause of issues working with SSO. If you still can't get the add-in to run successfully, see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).

## Stop the add-in

# [Visual Studio 2019](#tab/vs2019)

Choose **Stop Debugging**, or press <kbd>Shift</kbd>+<kbd>F5</kbd>.

# [Yo Office](#tab/yooffice)

Run `npm stop` from the command line.

---

## See also

[Using claims to reliably identify a user (Subject and Object ID)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)
