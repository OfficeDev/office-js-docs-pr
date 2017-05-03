# Create a NodeJS Office Add-in that uses Single Sign-on (preview)

Users can sign into Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign-on a second time. For an overview, see [Single Sign-on to Office, your Office Web Add-in, and Microsoft Graph (preview)](..\docs\develop\sso-in-office-add-ins.md) .

This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with NodeJS and express. 

> Note: For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses Single Sign-on](..\docs\develop\create-sso-office-add-ins-aspnet.md) .

## Prerequisites

* [Node and npm](https://nodejs.org/en/), version 6.9.4 or later.
* [Git Bash](https://git-scm.com/downloads) (Or another git client.)
* TypeScript version 2.2.2 or later.
* Office 2016, Version 1704,  build 8027.nnnn or later. (The Office 365 subscription version, sometimes called “Click to Run”.)  You many need to be an Office Insider to obtain this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1) .

## Setup the starter project

1. Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso). 


    > Note: There are two versions of the sample. 
    > 
    > * The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it. 
    > * The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.

1. Open a Git bash console in the **Before** folder.

2. Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.

3. Enter `npm run build ` in the console to build the project. 
     > Note: You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample missing some code that will be added later.

## Register the add-in with Azure AD V2 endpoint

1. Navigate to [https://apps.dev.microsoft.com/?test=build2017](https://apps.dev.microsoft.com/?test=build2017) . 

1. Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com

1. Click **Add an app**.

1. When prompted, use “Office-Add-in-NodeJS-SSO” as the app name, and then press **Create application**.

1. When the configuration page for the app opens, copy the **Application Id** and save it. You will use it in a later procedure. 

    > Note: This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You will need it in a later procedure. Then close the dialog.

1. In the **Platforms** section, click **Add Platform**. 

1. In the dialog that opens, select **Web API**.

1. An **Application ID URI** has been generated of the form “api://{App ID GUID}”. Replace the GUID with “localhost:3000”. The entire ID should read `api://localhost:3000`. (The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match. It should read `api://localhost:3000/access_as_user`.)

1. This step and the next one give the Office host application access to your add-in. In the **Pre-authorized applications** section, there is an empty **Application ID** box. Enter the following ID in the box (this is the ID of Microsoft Office):  `d3590ed6-52b3-4102-aeff-aad2292ab01c`.

1. Open the **Scope** drop down beside the **Application ID** and check the box for `api://localhost:3000/access_as_user`.

1. Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.

1. In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://localhost:3000`. 

    > Note: As of this writing, the **Web API** platform sometimes disappears from the **Platforms** section, particularly if the page is refreshed after the **Web** platform is added *and the registration page is saved*. For reassurance that your **Web API** platform is still part of the registration, click the **Edit Application Manifest** button near the bottom of the page. You should see the `api://localhost:3000` string in the **identifierUris** property of the manifest. There will also be a **oauth2Permissions** property whose **value** subproperty has the value `access_as_user`.

1. Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.

1. In the dialog, check the boxes for the following permissions (some may already be checked by default): 
    * Files.Read.All
    * profile


1. Click **OK** at the bottom of the dialog.

1. Click **Save** at the bottom of the registration page.

## Grant admin consent to the add-in

1. In the following string, replace the placeholder “{application_ID}” with the Application ID that you copied when you registered your add-in.

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Paste the resulting URL into a browser address bar and navigate to it.

1. When prompted, sign-in with the admin credentials to your Office 365 tenancy.

1. You are then prompted to grant permission for your add-in to access your Microsoft Graph data. Click **Accept**. 

1. The browser window/tab is then redirected to the **Redirect URL** that you specified when you registered the add-in; so, if the add-in is running, you the home page of the add-in opens in the browser. If the add-in is not running, you will get an error saying that the resource at localhost:3000 cannot be found or opened. *But the fact that the redirection was attempted means that the admin consent process completed successfully*. So regardless of whether the home page opened or you got the error, you can go on to the next step.

2. In the browser's address bar you'll see a "tenant" query parameter with a GUID value. This is the ID of your Office 365 tenancy. Copy and save this value. You will use it in a later step.

3. Close the window/tab.

## Configure the add-in

1. In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.

2. For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in. When you are done, there should just be a GUID in single quotation marks. There should not be any "{}" characters.

3. For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.

4. For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)
  
3. In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*. Replace this with the Office 365 tenancy ID that you saved at the end of the last procedure. If for any reason, you didn't get the ID earlier, use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it. When you are done, the `issuer` property value should look something like this:

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

    >Note: Leave the other parameters in the `AuthModule` constructor unchanged.

1. Save and close the file.

1. In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.

1. Scroll to the bottom of the file.

1. Just above the end `</VersionOverrides>` tag, you will find the following markup:

    ```
    <WebApplicationId>{application_GUID here}</WebApplicationId>
    <WebApplicationResource>api://localhost:3000<WebApplicationResource>
    <WebApplicationScopes>
        <WebApplicationScope>profile</WebApplicationScope>
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

1. Open the program.js file in the **public** folder. It already has some code in it:

    * An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.
    * A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.

1. Below the assignment to `Office.initialize`, add the code below. Note the following about this code: 

     * The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token. 
     * If no user is signed into Office, the Office host will prompt the user to sign in. 
     * The options parameter sets `forceConsent` to false, so the user will not be prompted to consent to giving the Office host access to your add-in.

    ```
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

    ```
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. Below the `getOneDriveFiles` method, add the following. This utility method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph. 

    ```
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

There are two server-side files that need to be modified. 
- The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.
- The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.

### Create a method to exchange tokens

1. Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:
    * The jwt parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.
    * The scopes parameter has a default value, but in this sample it will be overridden by the calling code.
    * The resource parameter is optional. It should not be used when the STS is the AAD V2 endpoint. The latter infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request. 
    

    ```
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // TODO2: Construct the parameters that will be sent in the body of the 
            //        HTTP Request to the STS that starts the "on behalf of" flow.
            // TODO3: Send the request to the STS.
            // TODO4: Process the response and persist the access token to resource.
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                        + ' ' + exception.message, 
                                        exception);
        }
    }
    ```

2. Replace TODO2 with the following code. About this code, note:
    * An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request. 
    * A resource property is added to the body if, and only if, a resource was passed to the method.

    ```
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

3. Replace TODO3 with the following code which sends the HTTP request to the token endpoint of the STS.

    ```
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Replace TODO4 with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.

    ```
    if (res.status !== 200) {
        const exception = await res.json();
        throw exception;
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

5. Save the file, but don't close it.

### Create a method to get access to the resource using the "on behalf of" flow

1. Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:
    * The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.
    * The method first checks the persistent storage for an access token to the resource that has not expired and is not going to in the next minute. It calls the method you created in the last section only if it needs to.

    ```
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

    ```
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Add the following method to bottom of the file. This method will handle any requests for the `onedriveitems` API.
    ```
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO5: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO6: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO7: Use the token to get data from Microsoft Graph.
        // TODO8: Send to the client only the data that it actually needs.
    })); 
    ```

4. Replace TODO5 with the following code which validates the access token received from the Office host application. The `verifyJWT` method is defined in the src\auth.ts file. It always validates the audience and the issuer. We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`. This is the only permisison to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf flow". 

    ```
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

5. Replace TODO6 with the following line. Note the following about this code:

    * The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2 endpoint which does not support a resource property.
    * The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive for Business.

    `const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['profile', 'Files.Read.All']);`

6. Replace TODO7 with the following line. Note the following about this code:

    * The MSGraphHelper class is defined in src\msgraph-helper.ts. 
    * We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. Replace TODO8 with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.

    ```
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
 2. Open a node terminal and navigate to the root folder of the project.
 3. In the terminal, enter **npm run build**. 
 4. Open a second node terminal and navigate to the root folder of the project.
 5. In the terminal, enter **npm run start**.

2. If you are using VS Code:
 3. Open the project in VS Code.
 4. Press CTRL-SHIFT-B to build the project.
 5. Press F5 to run the project in a debugging session.


## Add the add-in to an Office document

1. Restart PowerPoint and open or create a presentation. 

2. On the **Developer** tab in PowerPoint, choose **My Add-ins**.

3. Select the **SHARED FOLDER** tab.

4. Choose **SSO NodeJS Sample**, and then select **OK**.

5. On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon. 

## Test the add-in

> Note: The preview version of the `getAccessTokenAsync` API only supports work or school (Office 365) identities. *If you are signed into Office with a personal identity (Microsoft Account), sign out before preceding.* To test the add-in, you must be either signed out entirely from Office, or signed in with a work or school account.

1. Make sure you have some files or folders in your OneDrive for Business account.

2. Click **Show Add-in** button to open the add-in.

2. The add-in opens with a Welcome page. Click the **Get my files from OneDrive** button.

2. If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.

3. If you are not signed into Office, a popup will open and prompt you to sign in. After you have completed the sign-in, the list of your files and folders will appear after a few seconds. *You do not press the button a second time.*


