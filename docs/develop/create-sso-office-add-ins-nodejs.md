---
title: Create a Node.js Office Add-in that uses single sign-on
description: 
ms.date: 11/20/2017 
---

# Create a Node.js Office Add-in that uses single sign-on (preview)

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express. 

> [!NOTE]
> For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).

## Prerequisites

* [Node and npm](https://nodejs.org/en/), version 6.9.4 or later

* [Git Bash](https://git-scm.com/downloads) (or another git client)

* TypeScript version 2.2.2 or later

* Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”)

  You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## Set up the starter project

1. Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso). 

    > [!NOTE]
    > There are two versions of the sample:  
    > * The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it. 
    > * The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.

2. Open a Git bash console in the **Before** folder.

3. Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.

4. Enter `npm run build ` in the console to build the project. 

    > [!NOTE]
    > You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.

## Register the add-in with Azure AD v2.0 endpoint

1. Navigate to [https://apps.dev.microsoft.com](https://apps.dev.microsoft.com) . 

1. Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com

1. Click **Add an app**.

1. When prompted, use “Office-Add-in-NodeJS-SSO” as the app name, and then press **Create application**.

1. When the configuration page for the app opens, copy the **Application Id** and save it. You will use it in a later procedure. 

    > [!NOTE]
    > This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You will need it in a later procedure. Then close the dialog.

1. In the **Platforms** section, click **Add Platform**. 

1. In the dialog that opens, select **Web API**.

1. An **Application ID URI** has been generated of the form “api://{App ID GUID}”. Insert the string “localhost:3000” between the double forward slashes and the GUID. The entire ID should read `api://localhost:3000/{App ID GUID}`. 

    > [!NOTE]
    > The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match. It should read `api://localhost:3000/{App ID GUID}/access_as_user`.

1. This step and the next one give the Office host application access to your add-in. In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized. Each time you enter one, a new empty textbox appears. (Enter only the GUID.)

    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online) 

1. Open the **Scope** dropdown beside each **Application ID** and check the box for `api://localhost:44355/{App ID GUID}/access_as_user`.

1. Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.

1. In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://localhost:3000`. 

    > [!NOTE]
    > As of this writing, the **Web API** platform sometimes disappears from the **Platforms** section, particularly if the page is refreshed after the **Web** platform is added *and the registration page is saved*. For reassurance that your **Web API** platform is still part of the registration, click the **Edit Application Manifest** button near the bottom of the page. You should see the `api://localhost:3000/{App ID GUID}` string in the **identifierUris** property of the manifest. There will also be a **oauth2Permissions** property whose **value** subproperty has the value `access_as_user`.

1. Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.

1. In the dialog box, check the boxes for the following permissions: 

    * Files.Read.All
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.

1. Click **OK** at the bottom of the dialog.

1. Click **Save** at the bottom of the registration page.

## Grant admin consent to the add-in

> [!NOTE]
> This procedure is only needed when you are developing the add-in. When your production add-in is deployed to the Office Store or an add-in catalog, users will individually trust it when they install it.

1. In the following string, replace the placeholder “{application_ID}” with the Application ID that you copied when you registered your add-in.

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Paste the resulting URL into a browser address bar and navigate to it.

1. When prompted, sign-in with the admin credentials to your Office 365 tenancy.

1. You are then prompted to grant permission for your add-in to access your Microsoft Graph data. Click **Accept**. 

1. The browser window/tab is then redirected to the **Redirect URL** that you specified when you registered the add-in; so, if the add-in is running, the home page of the add-in opens in the browser. If the add-in is not running, you will get an error saying that the resource at localhost:3000 cannot be found or opened. *But the fact that the redirection was attempted means that the admin consent process completed successfully*. So regardless of whether the home page opened or you got the error, you can go on to the next step.

2. In the browser's address bar you'll see a "tenant" query parameter with a GUID value. This is the ID of your Office 365 tenancy. Copy and save this value. You will use it in a later step.

3. Close the window/tab.

## Configure the add-in

1. In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.

2. For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in. When you are done, there should just be a GUID in single quotation marks. There should not be any "{}" characters.

3. For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.

4. For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)
  
3. In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*. Replace this with the Office 365 tenancy ID that you saved at the end of the last procedure. If for any reason, you didn't get the ID earlier, use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it. When you are done, the `issuer` property value should look something like this:

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. Leave the other parameters in the `AuthModule` constructor unchanged. Save and close the file.

1. In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.

1. Scroll to the bottom of the file.

1. Just above the end `</VersionOverrides>` tag, you will find the following markup:

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in. (The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.

    > [!NOTE]
    > * The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.
    > * The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through the Office Store.

1. Save and close the file.

## Code the client side

1. Open the program.js file in the **public** folder. It already has some code in it:

    * An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.
    * A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.

1. Below the assignment to `Office.initialize`, add the code below. Note the following about this code: 

    * The `getDataWithoutAuthChallenge` function is called in a first attempt to use the on-behalf-of flow. The assumption is that single factor authentication is all that is needed. You'll add code in a later step to handle the case where multi-factor authenticatiion is needed.
    * The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token. 
     * If no user is signed into Office, the Office host will prompt the user to sign in. 
     * The options parameter sets `forceConsent` to false, so the user will not be prompted to consent to giving the Office host access to your add-in.

    ```javascript
    function getOneDriveItems() {
        getDataWithoutAuthChallenge();
    }	
    
    function getDataWithoutAuthChallenge() {       
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
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

1. Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/onedriveitems” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.

    ```javascript
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. Below the `getOneDriveFiles` method, add the following. This utility method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph. 

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            TODO2: Display data and handle demand for multi-factor authentication.
        })
        .fail(function (result) {
            console.log(result.error);
       });
    }
    ```

1. Replace TODO2 with the following code. About this code, note:
    * If the Microsoft Graph target requests addtional authentication factor(s), the result will not be data. It will be a Claims JSON telling AAD what addtional factors the user must provide. In that case, the client must start a new sign-on that passes this Claims string to AAD so that AAD will provide the needed prompts.
    * If the result is the Claims JSON, then it will contain the string "capolids".
    * You will create the `getDataUsingAuthChallenge` function in a latter step.

    ```javascript
    if (result[0].indexOf('capolids') !== -1) {                
        result[0] = JSON.parse(result[0])
        getDataUsingAuthChallenge(result[0]);
    } else {  
        showResult(result);
    }
    ```

1. Add the following function to the file just below the `getData` function. About this function, note:
    * The function is used when AAD has requested additional authentication factor(s). 
    * The function triggers a second sign-on in which the user will be prompted to provide additional authentication factor(s). 
    * The `authChallenge` option contains a string that tells AAD what factor(s) it should prompt for. The Office host passes this string to AAD when it requests the add-in token to your add-in.

    ```javascript
    function getDataUsingAuthChallenge(authChallengeString) {       
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/onedriveitems", accessToken);
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

1. Save and close the file.

## Code the server side

There are two server-side files that need to be modified. 
- The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.
- The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.

### Create a method to exchange tokens

1. Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:

    * The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.
    * The scopes parameter has a default value, but in this sample it will be overridden by the calling code.
    * The resource parameter is optional. It should not be used when the STS is the AAD V2 endpoint. The latter infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request. 
    
        ```javascript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. Replace TODO3 with the following code. About this code, note:
    * An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request. 
    * A resource property is added to the body if, and only if, a resource was passed to the method.

        ```javascript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            } 
        ```

3. Replace TODO4 with the following code which sends the HTTP request to the token endpoint of the STS.

    ```javascript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Replace TODO5 with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.

    ```javascript
    if (res.status !== 200) {
        TODO6: Handle failure and the case where AAD asks for additional
               authentication factors.
    }
    const json = await res.json();
    // Persist the token and it's expiration time.
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

5. Replace the TODO6 with the following code. About this code, note:

    * There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response that has a `Claims` property. 
    * This `Claims` value needs to be passed back to the client which should initiate a second sign-on for the user and include the `Claims` value in the call to the AAD. AAD will prompt the user to provide the addtional factor(s).
    * As a precaution, the code clears the cache of any access tokens that were obtained when the user logged in with only a password.  

    ```javascript
    const exception = await res.json();
    // Check if AAD is the STS.
    if (this.stsDomain === 'https://login.microsoftonline.com') {
        if (JSON.stringify(exception.claims)) {                       
            ServerStorage.clear();
            return JSON.stringify(exception.claims);    
        } else {                    
            throw exception;
        }
    }
    else {                    
        throw exception;
    }
    ```

5. Save the file, but don't close it.

### Create a method to get access to the resource using the "on behalf of" flow

1. Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:

    * The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.
    * The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.

    ```javascript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. Save and close the file.

### Create the endpoints that will serve the add-in's home page and data

1. Open the src\server.ts file. 

2. Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.

    ```javascript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Add the following method to bottom of the file. This method will handle any requests for the `onedriveitems` API.
    ```javascript
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Send to the client only the data that it actually needs.
    })); 
    ```

4. Replace TODO7 with the following code which validates the access token received from the Office host application. The `verifyJWT` method is defined in the src\auth.ts file. It always validates the audience and the issuer. We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`. This is the only permisison to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf flow". 

    ```javascript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.

5. Replace TODO8 with the following code. Note the following about this code:

    * The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.
    * The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive. (The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)
    * If the response is a string containing 'capolids", then this is a claims message from AAD that multi-factor auth is required. The message is passed to the client, which uses it to start a second sign-on. The string tells AAD what additional authentication factor(s) it should prompt the user to provide.

    ```javascript
    let graphToken = null;
    const tokenAcquisitionResponse = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    if (tokenAcquisitionResponse.includes('capolids')) {
        const claims: string[] = [];
        claims.push(tokenAcquisitionResponse);
        return res.json(claims);
    } else {
        // The response is the token to Microsoft Graph itself. Rename it so remaining code
        // is self-documenting.
        graphToken = tokenAcquisitionResponse;
    }
    ```

6. Replace TODO9 with the following line. Note the following about this code:

    * The MSGraphHelper class is defined in src\msgraph-helper.ts. 
    * We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.

    ```javascript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

7. Replace TODO10 with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.

    ```javascript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. Save and close the file.

## Deploy the add-in

Now you need to let Office know where to find the add-in.

1. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx).

2. Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.

3. Launch PowerPoint and open a document.

4. Choose the **File** tab, and then choose **Options**.

5. Choose **Trust Center**, and then choose the **Trust Center Settings** button.

6. Choose **Trusted Add-ins Catalogs**.

7. In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.

8. Select the **Show in Menu** check box, and then choose **OK**.

9. A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.

## Build and run the project

There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.

1. If you are not using Visual Studio Code: 
 1. Open a node terminal and navigate to the root folder of the project.
 2. In the terminal, enter **npm run build**. 
 3. Open a second node terminal and navigate to the root folder of the project.
 4. In the terminal, enter **npm run start**.

2. If you are using VS Code:
 1. Open the project in VS Code.
 2. Press CTRL-SHIFT-B to build the project.
 3. Press F5 to run the project in a debugging session.


## Add the add-in to an Office document

1. Restart PowerPoint and open or create a presentation. 

2. On the **Developer** tab in PowerPoint, choose **My Add-ins**.

3. Select the **SHARED FOLDER** tab.

4. Choose **SSO NodeJS Sample**, and then select **OK**.

5. On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon. 

## Test the add-in

1. Ensure that you have some files in your OneDrive so that you can verify the results.

2. Click **Show Add-in** button to open the add-in.

2. The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.

2. If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.

3. If you are not signed into Office, a popup will open and prompt you to sign in. After you have completed the sign-in, the list of your files and folders will appear after a few seconds. *You do not press the button a second time.*

> [!NOTE]
> If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.
