---
title: Create a Node.js Office Add-in that uses single sign-on
description: 'Learn how to create a Node.js-based add-in that uses Office Single Sign-on'
ms.date: 01/13/2020
localization_priority: Priority
---

# Create a Node.js Office Add-in that uses single sign-on (preview)

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express. For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).

> [!NOTE]
> As an alternative to completing the steps described in this article, you can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in. The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO. For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).

## Prerequisites

* [Node and npm](https://nodejs.org/), version 10.15.0 or later.

* [Git Bash](https://git-scm.com/downloads) (or another git client)

* TypeScript, version 3.6.2 or later

* Office 365 (the subscription version of Office) account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365. You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.

* A code editor. We recommend Visual Studio Code.

* At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.

* A Microsoft Azure subscription. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Set up the starter project

1. Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).

    > [!NOTE]
    > There are three versions of the sample:  
    > * The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
    > * The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.
    > * The **SSOAutoSetup** version is a completed sample that automates most of the steps to register the add-in with Azure AD and configure it. Use this version if you want to see a working add-in with SSO quickly. Just follow the steps in the Readme of the folder. We recommend that at some point you go through the manual registration and setup steps in this article to better understand the relationship between Azure AD and an add-in. 


1. Open a command prompt in the **Before** folder.

1. Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.

1. Run the command `npm run install-dev-certs`. Select **Yes** to the prompt to install the certificate.

## Register the add-in with Azure AD v2.0 endpoint

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.

1. Sign in with the ***admin*** credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    * Set **Name** to `Office-Add-in-NodeJS-SSO`.
    * Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.
    * Set **Redirect URI** to` https://localhost:44355/dialog.html`.
    * Choose **Register**.

1. On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**. You'll use both of them in later procedures.

    > [!NOTE]
    > This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. Select **Authentication** under **Manage**. In the **Implict grant** section, enable the checkboxes for both **Access token** and **ID token**. The sample has a fallback authorization system that is invoked when SSO is not available. This system uses the Implicit Flow.

1. Select **Save** at the top of the form.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**. *Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.

1. Select **Expose an API** under **Manage**. Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**. Insert `localhost:44355/` (note the forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`. 

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope** name.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user. Suggestions:

    - **Admin consent title**: Office can act as the user.
    - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.
    - **User consent title**: Office can act as you.
    - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope** .

    > [!NOTE]
    > The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office on the web)

    For each ID, take these steps:

    a. Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.

    b. Select **Add application**.

1. Select **API permissions** under **Manage** and select **Add a permission**. On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search for the permissions your add-in needs. Select the following. Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.

    * Files.Read.All
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. Select the check box for each permission as it appears. After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.

1. On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.

## Configure the add-in

1. Open the `\Begin` folder in the cloned project in your code editor.

1. Open the `.ENV` file and use the values that you copied earlier. Set the **CLIENT_ID** to your **Application (client) ID**, and set the **CLIENT_SECRET** to your client secret. The values should **not** be in quotation marks. When you are done, the file should be similar to the following: 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. Open the `\public\javascripts\fallbackAuthDialog.js` file. In the `msalConfig` declaration, replace the placeholder $application_GUID here$ with the Application ID that you copied when you registered your add-in. The value should be in quotation marks.

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup:

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Replace the placeholder "$application_GUID here$" *in both places* in the markup with the Application ID that you copied when you registered your add-in. The "$" symbols are not part of the ID, so do not include them. This is the same ID you used in for the ClientID and Audience in the web.config.

	> [!NOTE]
    > The **Resource** value is the **Application ID URI** you set when you registered the add-in. The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.

## Code the client-side

### Create the SSO logic

1. In your code editor, open the file `public\javascripts\ssoAuthES6.js`. It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.

	> [!NOTE]
    > As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API. When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will run in Internet Explorer 11. 

1. Add the following code below the Office.onReady method:

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exhange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. Replace `TODO 1` with the following code. About this code, note:

    - `OfficeRuntime.auth.getAccessToken` instructs Office to get a bootstrap token from Azure AD. A bootstrap token is similar to an ID token, but it has a `scp` (scope) property with the value `access-as-user`. This kind of token can be exchanged by a web application for an access token to Microsoft Graph.
    - Setting the `allowSignInPrompt`option to true means that if no user is currently signed into Office, then Office will open a popup sign-in prompt.
    - Setting the `forMSGraphAccess` option to true signals to Office that the add-in intends to use the bootstrap token to get an access token to Micrsoft Graph, instead of just using it as an ID token. If the tenant administrator has not granted consent to the add-in's access to Microsoft Graph, then `OfficeRuntime.auth.getAccessToken` returns error **13012**. The add-in can respond by falling back to an alternative system of authorization, which is necessary because Office can prompt only for consent to the user's Azure AD profile, not to any Microsoft Graph scopes. The fallback authorization system requires the user to sign in again and the user *can* be prompted to consent to Micrsoft Graph scopes. So, the `forMSGraphAccess` option ensures that the add-in won't make a token exchange that will fail due to lack of consent. (Since you granted administrator consent in an earlier step, this scenario won't happen for this add-in. But the option is included here anyway to illustrate a best practice.)

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
    ```

1. Replace `TODO 2` with the following code. You'll create the `getGraphToken` method in a later step.

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. Replace `TODO 3` with the following. About this code, note: 

    - If the Office 365 tenant has been configured to require multifactor authentication, then the `exchangeResponse` will include a `claims` property with information about the additional required factors. In that case, `OfficeRuntime.auth.getAccessToken` should be called again with the `authChallenge` option set to the value of the claims property. This tells AAD to prompt the user for all required forms of authentication.

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. Replace `TODO 4` with the following. About this code, note: 

    - You'll create the `handleAADErrors` method in a later step. Azure AD errors are returned to the client as HTTP code 200 Responses. They do not throw errors, so they do not trigger the `catch` block of the `getGraphData` method.
    - You'll create the `makeGraphApiCall` method in a later step. It makes an AJAX call to the MS Graph endpoint. Errors are caught in the `.fail` callback of that call, not in the `catch` block of the `getGraphData` method.

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. Replace `TODO 5` with the following

    - Errors from the call of `getAccessToken` will have a `code` property with an error number, typically in the 13xxx range. You'll create the `handleClientSideErrors` method in a later step.
    - The `showMessage` method displays text on the task pane.

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. Below the `getGraphData` method, add the following function. Note that `/auth` is a server-side Express route that exhanges the bootstrap token with Azure AD for an access token to Microsoft Graph.

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. Below the `getGraphToken` method, add the following function. Note that `error.code` is a number, usually in the range 13xxx.

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```
1. Replace `TODO 6` with the following code. 
For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the Web.
        showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. Replace `TODO 7` with the following code. For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). The function `dialogFallback` invokes the alternative system of authorization. In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. Below the `handleClientSideErrors` function, add the following function. 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. On rare occasions the bootstrap token that Office has cached is unexpired when Office validates it, but expires by the time it reaches Azure AD for exchange. Azure AD will respond with error **AADSTS500133**. In this case, the add-in should simply recursively call `getGraphData`. Since the cached bootstrap token is now expired, Office will get a new one from Azure AD. So replace `TODO 8` with the following. 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)       
    {
        getGraphData();
    }
    ```

1. To ensure that the add-in doesn't enter an infinite loop of calls to `getGraphData`, the add-in should keep track of how many times `getGraphData` has been called and be sure that is not called recursively called more than once. So, create a counter variable in a scope that is global to the `handleAADErrors` and `getGraphData` functions. A good place for global variables is just below the `Office.onReady` method call.

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. Change the `if` structure in the `handleAADErrors` method so that it:

    - Increments the counter just before it calls `getGraphData`.
    - Tests to ensure that `getGraphData` has not already been called a second time. 

    So the final version of the `if` structure should look like the following:

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Replace `TODO 9` with the following. 

    ```javascript
    else {                
        dialogFallback();
    }
    ```

1. Save and close the file.

### Get the data and add it to the Office document

1. In the `public\javascripts` folder, create a new file named `data.js`.

1. Add the following function to the file. This is the function that is called by the `getGraphData` function when it has acquired an access token to Microsoft Graph. 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. Replace `TODO 10` with the following. About this code, note: 

    - This object is the parameter to the `$.ajax` method.
    - The `/getuserdata` is an Express route on the add-in's server that you create in a later step. It will call a Microsoft Graph endpoint and include the access token in its call. 

    ```javascript
    {
        type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. Replace `TODO11` with the following. About this code, note:

    - The `writeFileNamesToOfficeDocument` will insert the data from Graph into the Office document. It is defined in the `public\javascripts\document.js` file. 
    - If `writeFileNamesToOfficeDocument` returns an error, it will begin with "Unable to add filenames to document."

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {        
        showMessage(error);
    });
    ```

1. Save and close the file.

## Code the server-side

### Create the auth router and the token exchange logic

1. Open the file `routes\authRoute.js` and add the following route function just below the `require` statements and above the `module.exports` statement. Note that the URL parameter of `router.get` is '/'. Since this route is being defined in a router that will handle all HTTP Requests for the URL '/auth', this route effectively handles all requests for '/auth'. The client-side `getGraphToken` function that you created earlier calls this route.  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exhange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. Replace `TODO 12` with the following code.

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. Replace `TODO 13` with the following code. About this code, note: 

    - This is the beginning of a long `else` block, but the closing `}` is not at the end yet because you will be adding more code to it. 
    - The `authorization` string is "Bearer " followed by the bootstrap token, so the first line of the `else` block is assigning the token to the `jwt`. ("JWT" stands for "JSON Web Token".)
    - The two `process.env.*` values are the constants that you assigned when you configured the add-in. 
    - The `requested_token_use` form parameter is set to 'on_behalf_of'. This tells Azure AD that the add-in is requesting an access token to Microsoft Graph using the On-Behalf-Of Flow. Azure will respond by validating that the bootstrap token, which is assigned to `assertion` form parameter, has a `scp` property that is set to `access-as-user`.
    - The `scope` form parameter is set to 'Files.Read.All' which is the only Microsoft Graph scope that the add-in needs.

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. Replace `TODO 14` with the following code, which completes the `else` block. About this code, note:

    - The const `tenant` is set to 'common' because you configured the add-in as multitenant when you registered it with Azure AD; specifically when you set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**. If you had instead chosen to support only accounts in the same Office 365 tenancy where the add-in is registered, then in this code `tenant` would be set to the GUID of the tenant. 
    - If the POST request does not error, then the response from Azure AD is converted to JSON and sent to the client. This JSON object has an `access_token` property to which Azure AD has assigned the access token to Microsoft Graph.

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();
            
            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. Save and close the file.

### Create the route that will fetch the data from Microsoft Graph

1. Open the file `app.js` in the root of the project. Just below the route for '/dialog.html', add the following route. This route is called by the `makeGraphApiCall` function that you created in an earlier step.

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. Replace `TODO 15` with the following. About this code, note:

    - The caller of this route, `makeGraphApiCall`, added the access token to Microsoft Graph to the HTTP Request as a header named "access_token".
    - The `getGraphData` function is defined in the `msgraph-helper.js` file. (This is not the same function as the client-side `getGraphData` function that you defined in the `ssoAuthES6.js` file.)
    - The last parameter, for `queryParamsSegment`, is hardcoded. If you reuse this code in a production add-in and any part of `queryParamsSegment` comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.
    - The code minimizes the data that must come from Microsoft Graph by specifying only the property we need ("name") and only the top 10 folder or file names.

    ```javascript
    const graphToken = req.get('access_token');    
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. Replace `TODO 16` with the following. About this code, note:

    - If Microsoft Graph returns an error, such as invalid or expired token, there will be a code property in the returned object set to a HTTP status (e.g., 401). The code relays the error to the client. It will be caught in the `.fail` callback of `makeGraphApiCall`.
    - Microsoft Graph data includes OData metadata and eTags that the add-in does not need, so the code constructs a new array containing only the file names to send to the client.

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. Save and close the file.

## Run the project

1. Ensure that you have some files in your OneDrive so that you can verify the results.

1. Open a command prompt in the root of the `\Complete` folder. 

1. Run the command `npm start`. 

1. You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it. The instructions depend on your platform. There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.

1. Click the **Get OneDrive File Names** button. If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document. (It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in. After you log in, the file and folder names appear.

> [!NOTE]
> If you were previously signed into Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get OneDrive File Names**.
