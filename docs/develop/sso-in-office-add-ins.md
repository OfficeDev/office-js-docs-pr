# Enable single sign-on for Office Add-ins

Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use SSO to do the following--without requiring the user to sign in a second time:

* Authorize the user to sign in to your add-in.
* Authorize the add-in to access [Microsoft Graph](https://developer.microsoft.com/graph/docs).

![An image showing the sign-in process for an add-in](../../images/OfficeHostTitleBarLogin.png)

>**Note:**
> The Single Sign-on API is currently supported for Word, Excel, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../../reference/requirement-sets/identity-api-requirement-sets.md).
> Single Sign-on is currently in preview for Outlook. If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy. Details are at [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)

For users, this makes running your add-in a smooth experience that involves at signing in only once. For developers, this means that your add-in can authenticate users and gain authorized access to the user’s data via Microsoft Graph with credentials that the user has already provided to the Office application.

### SSO add-in architecture

In addition to hosting the pages and JavaScript of the web application, the add-in must also host, at the same [fully qualified domain name](https://msdn.microsoft.com/en-us/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly), one or more web APIs that will get an access token to Microsoft Graph and make requests to it.

The add-in manifest contains markup that specifies how the add-in is registered in the Azure Active Directory (Azure AD) v2.0 endpoint, and it specifies any permissions to Microsoft Graph that the add-in needs.

### How it works at runtime

The following diagram shows how the SSO process works.
<!-- Minor fixes to the text in the diagram - change V2 to v2.0, and change "(e.g. Word, Excel, etc.)" to "(for example, Word, Excel)". -->
![A diagram that shows the SSO process](../../images/SSOOverviewDiagram.png)

1. In the add-in, JavaScript calls a new Office.js API `getAccessTokenAsync`. This tells the Office host application to obtain an access token to the add-in. (Hereafter, this is called the **add-in token**.)
1. If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.
1.  If this is the first time the current user has used your add-in, he or she is prompted to consent.
1. The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.
1. Azure AD sends the add-in token to the Office host application.
1. The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.
1. JavaScript in the add-in makes an HTTP request to a web API that is hosted at the same fully-qualified domain as the add-in, and it includes the **add-in token** as authorization proof.  
1. Server-side code validates the incoming **add-in token**.
1. Server-side code uses the “on behalf of” flow (defined at [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) and the [daemon or server application to web API Azure scenario](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) to obtain an access token for Microsoft Graph (hereafter, called the **MSG token**) in exchange for the add-in token.
1. Azure AD returns the **MSG token** (and a refresh token, if the add-in requests *offline_access* permission) to the add-in.
1. Server-side code caches the **MSG token(s)**.
1. Server-side code makes requests to Microsoft Graph and includes the **MSG token**.
1. Microsoft Graph returns data to the add-in, which can pass it on to the add-in’s UI.
1. When the MSG token expires, the server-side code can use its refresh token to get a new **MSG token**.

### Develop an SSO add-in

This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:

* [Create a Node.js Office Add-in that uses single sign-on](../../docs/develop/create-sso-office-add-ins-nodejs.md)
* [Create an ASP.NET Office Add-in that uses single sign-on](../../docs/develop/create-sso-office-add-ins-aspnet.md)

#### Create the service application

Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:

* Get a client ID and name for the add-in.
* Specify the permissions that your add-in needs to Microsoft Graph.
* Grant the Office host application trust to the add-in.
* Preauthorize the Office host application to the add-in with the default permission *access_as_user*.

#### Configure the add-in

Add new markup to the add-in manifest:

* **WebApplicationInfo** - The parent of the following elements.
* **Id** - The client ID of the add-in.
* **Resource** - The URL of the add-in.
* **Scopes** - The parent of one or more **Scope** elements.
* **Scope** - Specifies a permission that the add-in needs to Microsoft Graph. For example, `User.Read`, `Mail.Read` or `offline_access`). For more information, see [Microsoft Graph permissions](https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference)

For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.

#### Add client-side code

Add JavaScript to the add-in to:

* Call `Office.context.auth.getAccessTokenAsync(myTokenHandler)`.
* Create a handler that passes the add-in token to the add-in’s server-side code. For example:

```js
function mytokenHandler(asyncResult) {
    // Passes asyncResult.value (which has the add-in access token)
    // to the add-in’s web API as an Authorization header.
}
```

#### When to call the method

If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.

If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*. There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever `getAccessTokenAsync` is called. So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.

#### Add server-side code

Create one or more Web API methods that get Microsoft Graph data. Depending on your language and framework, libraries might be available that will simplify the code you have to write. Your server-side code should do the following:

* Validate the add-in token that is received from the token handler you created earlier.
* Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).
* Cache the returned MSG token.
* Get data from Microsoft Graph by using the MSG token.
