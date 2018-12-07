---
title: Create a Node.js Office Add-in that uses single sign-on
description: ''
ms.date: 12/7/2018
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

  You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).

## Set up the starter project

1. Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso). 

    > [!NOTE]
    > There are three versions of the sample:  
    > * The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it. 
    > * The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.
    > * The **Completed Multitenant** version is a completed sample that supports multitenancy. Explore this sample if you intend to support Microsoft accounts from different domains with SSO.
    >
    > _Regardless of which version you use, you will need to trust a certificate for the localhost. See the "IMPORTANT" note in the Readme of the repo._

2. Open a Git bash console in the **Before** folder.

3. Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.

4. Enter `npm run build ` in the console to build the project. 

    > [!NOTE]
    > You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.

## Register the add-in with Azure AD v2.0 endpoint

The following instruction are written generically so they can be used in multiple places. For this article do the following:
- Replace the placeholder **$ADD-IN-NAME$** with `“Office-Add-in-NodeJS-SSO`.
- Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.
- When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions. Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.
    * Files.Read.All
    * profile

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## Grant administrator consent to the add-in

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## Configure the add-in

1. In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.

2. For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in. When you are done, there should just be a GUID in single quotation marks. There should not be any "{}" characters.

3. For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.

4. For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)
  
3. In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*. Replace this with the Office 365 tenancy ID. Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it. When you are done, the `issuer` property value should look something like this:

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
    > * The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.

1. Save and close the file.

## Code the client side

1. Open the program.js file in the **public** folder. It already has some code in it:

    * An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.
    * A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.
    * A `logErrors` method that will log to console errors that are not intended for the end user.

11. Below the assignment to `Office.initialize`, add the code below. Note the following about this code:

    * The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token. 
    * You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

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
            //        nor Microsoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user triggered an operation that calls `getAccessTokenAsync` 
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

1. Replace `TODO4` with the following code. Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft Account. Ask the user to sign-out and then in again with a supported account type.

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
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

1. Replace `TODO7` with the following code. Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. Replace `TODO8` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
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
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. Replace `TODO10` with the following code. Note about this code:

    * There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property. 
    * The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication. 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. Replace `TODO11` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:

    * Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions. 
    * The add-in should get a new token with the `forceConsent` option set to `true`.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:

    * Error 70011 means that an invalid scope (permission) has been requested. The add-in should report the error.
    * The code logs any other error with an AAD error number.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:

    * Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.
    * The add-in should report the error.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:

    * It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.
    * In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method. But it should do this only once. If it happens again, it should just log the error.
    * The code logs the error if it happens twice in succession.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }        
    }
    ```

1. Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*.

    ```javascript
    else {
        logError(result);
    }
    ```

## Code the server side

There are two server-side files that need to be modified. 
- The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.
- The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.

### Create a method to exchange tokens

1. Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:

    * The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.
    * The scopes parameter has a default value, but in this sample it will be overridden by the calling code.
    * The resource parameter is optional. It should not be used when the [Secure Token Service (STS)](https://docs.microsoft.com/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) is the AAD V 2.0 endpoint. The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request. 
    * Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client. Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. Replace `TODO3` with the following code. About this code, note:
    * An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request. 
    * A resource property is added to the body if, and only if, a resource was passed to the method.

        ```typescript
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

3. Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Replace `TODO5` with the following code. Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client. Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. Save the file, but don't close it.

### Create a method to get access to the resource using the "on behalf of" flow

1. Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:

    * The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.
    * The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.

    ```typescript
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

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Add the following method to bottom of the file. This method will handle any requests for the `values` API.
    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. Replace `TODO7` with the following code which validates the access token received from the Office host application. The `verifyJWT` method is defined in the src\auth.ts file. It always validates the audience and the issuer. We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`. This is the only permission to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf" flow. 

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.

5. Replace `TODO8` with the following code. Note the following about this code:

    * The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.
    * The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive. (The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. Replace `TODO9` with the following line. Note the following about this code:

    * The MSGraphHelper class is defined in src\msgraph-helper.ts. 
    * We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

7. Replace `TODO10` with the following code. Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token. It is very unlikely that this would ever happen since the token persisting logic should prevent it. (See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name. (See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. Save and close the file.

### Add response handling to the ODataHelper

1. Open the file src\odata-helper.ts. The file is almost complete. What's missing is the body of the callback to the handler for the request "end" event. Replace the `TODO` with the following code. About this code note:

    * The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired. But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered. So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  Replace `TODO2` with the following code. Note about this code:

    * An error response from an OData source will always have a statusCode and usually a statusMessage. Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.
    * The Promise object is resolved, not rejected. The `https.get` runs when a web service calls an OData endpoint server-to-server. But that call comes in the context of a call from a client to a web API in the web service. The "outer" request from the client to the web service never completes if this "inner" request is rejected. Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;
    
    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. Save and close the file.

## Deploy the add-in

Now you need to let Office know where to find the add-in.

1. Create a network share, or [share a folder to the network](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).

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

1. If the **Developer** tab is not visible on the ribbon, enable it with the following steps:
 1. Navigate to **File** | **Options** | **Customize Ribbon**.
 2. Click the check box to enable **Developer** in the tree of control names on the right of the **Customize Ribbon** page.
 3. Press **OK**.

2. On the **Developer** tab in PowerPoint, choose **My Add-ins**.

3. Select the **SHARED FOLDER** tab.

4. Choose **SSO NodeJS Sample**, and then select **OK**.

5. On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon. 

## Test the add-in

1. Ensure that you have some files in your OneDrive so that you can verify the results.

2. Click **Show Add-in** button to open the add-in.

2. The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.

2. If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.

3. If you are not signed into Office, a popup will open and prompt you to sign in. After you have completed the sign-in, the list of your files and folders will appear after a few seconds. *You should not press the button a second time.*

> [!NOTE]
> If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.
