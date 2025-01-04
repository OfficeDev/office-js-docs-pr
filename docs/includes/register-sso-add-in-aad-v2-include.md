## Register the add-in with Microsoft identity platform

You need to create an app registration in Azure that represents your web server. This enables authentication support so that proper access tokens can be issued to the client code in JavaScript. This registration supports both SSO in the client, and fallback authentication using the Microsoft Authentication Library (MSAL).

1. Sign in to the [Azure portal](https://portal.azure.com/) with the ***admin*** credentials to your Microsoft 365 tenancy. For example, **MyName@contoso.onmicrosoft.com**.
1. Select **App registrations**. If you don't see the icon, search for "app registration" in the search bar.

    :::image type="content" source="../images/azure-portal-select-app-registration.png" alt-text="The Azure portal home page.":::

    The **App registrations** page appears.

1. Select **New registration**.

    :::image type="content" source="../images/azure-portal-select-new-registration.png" alt-text="New registration on the App registrations pane.":::

    The **Register an application** page appears.

1. On the **Register an application** page, set the values as follows.

    * Set **Name** to `<add-in-name>`.
    * Set **Supported account types** to **Accounts in any organizational directory (any Azure AD directory - multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    * Set **Redirect URI** to use the platform **Single-page application (SPA)** and the URI to `https://<fully-qualified-domain-name>/dialog.html`.

    :::image type="content" source="../images/azure-portal-register-an-application.png" alt-text="Register an application pane with name and supported account completed.":::

1. Select **Register**. A message is displayed stating that the application registration was created.

    :::image type="content" source="../images/azure-portal-application-created-message.png" alt-text="Message stating that the application registration was created.":::

1. Copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**. You'll use both of them in later procedures.

    :::image type="content" source="../images/azure-portal-copy-client-directory-ids.png" alt-text="App registration pane for Contoso displaying the client ID and directory ID.":::

## Add a client secret

Sometimes called an _application password_, a client secret is a string value your app can use in place of a certificate to identity itself.

1. From the left pane, select **Certificates & secrets**. Then on the **Client secrets** tab, select **New client secret**.

    :::image type="content" source="../images/azure-portal-create-new-client-secret.png" alt-text="The Certificates & secrets pane.":::

    The **Add a client secret** pane appears.

1. Add a description for your client secret.
1. Select an expiration for the secret or specify a custom lifetime.
    * Client secret lifetime is limited to two years (24 months) or less. You can't specify a custom lifetime longer than 24 months.
    * Microsoft recommends that you set an expiration value of less than 12 months.

    :::image type="content" source="../images/azure-portal-client-secret-description.png" alt-text="Add a client secret pane with description and expires completed.":::

1. Select **Add**. The new secret is created and the value is temporarily displayed.

> [!IMPORTANT]
> _Record the secret's value_ for use in your client application code. This secret value is _never displayed again_ after you leave this pane.

## Expose a web API

1. From the left pane, select **Expose an API**.

    The **Expose an API** pane appears.

    :::image type="content" source="../images/azure-portal-expose-an-api.png" alt-text="An app registration's Expose an API pane.":::

1. Select **Set** to generate an application ID URI.

    :::image type="content" source="../images/azure-portal-set-api-uri.png" alt-text="Set button in the app registration's Expose an API pane.":::

    The section for setting the application ID URI appears with a generated Application ID URI in the form `api://<app-id>`.

1. Update the application ID URI to `api://<fully-qualified-domain-name>/<app-id>`.

    :::image type="content" source="../images/azure-portal-app-id-uri-details.png" alt-text="Edit the App ID URI pane with localhost port set to 44355.":::

    * The **Application ID URI** is pre-filled with app ID (GUID) in the format `api://<app-id>`.
    * The application ID URI format should be: `api://<fully-qualified-domain-name>/<app-id>`
    * Insert the `fully-qualified-domain-name` between `api://` and `<app-id>` (which is a GUID). For example, `api://contoso.com/<app-id>`.
    * If you're using localhost, then the format should be `api://localhost:<port>/<app-id>`. For example, `api://localhost:3000/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    For additional application ID URI details, see [Application manifest identifierUris attribute](/azure/active-directory/develop/reference-app-manifest#identifieruris-attribute).

    > [!NOTE]
    > If you get an error saying that the domain is already owned but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step. (This error can also occur if you are not signed in with credentials of an admin in the Microsoft 365 tenancy. See step 2. Sign out and sign in again with admin credentials and repeat the process from step 3.)

## Add a scope

1. On the **Expose an API** page, select **Add a scope**.

    :::image type="content" source="../images/azure-portal-add-a-scope.png" alt-text="Select Add a scope button.":::

    The **Add a scope** pane opens.

1. In the **Add a scope** pane, specify the scope's attributes. The following table shows example values for and Outlook add-in requiring the `profile`, `openid`, `Files.ReadWrite`, and `Mail.Read` permissions. Modify the text to match the permissions your add-in needs.

    | Field | Description | Values |
    |--|--|--|
    | **Scope name** | The name of your scope. A common scope naming convention is `resource.operation.constraint`. | For SSO this must be set to `access_as_user`. |
    | **Who can consent** |  Determines if admin consent is required or if users can consent without an admin approval. | For learning SSO and samples, we recommend you set this to **Admins and users**. <br><br>Select **Admins only** for higher-privileged permissions.|
    | **Admin consent display name** | A short description of the scope's purpose visible to admins only. | `Read/write permissions to user files. Read permissions to user mail and profiles.` |
    | **Admin consent description** | A more detailed description of the permission granted by the scope that only admins see. | `Allow Office to have read/write permissions to all user files and read permissions to all user mail. Office can call the app's web APIs as the current user.` |
    | **User consent display name** | A short description of the scope's purpose. Shown to users only if you set **Who can consent** to **Admins and users**. | `Read/write permissions to your files. Read permissions to your mail and profile.` |
    | **User consent description** | A more detailed description of the permission granted by the scope. Shown to users only if you set **Who can consent** to **Admins and users**. | `Allow Office to have read/write permissions to your files, and read permissions to your mail and profile.` |

1. Set the **State** to **Enabled**, and then select **Add scope**.

    :::image type="content" source="../images/azure-portal-enable-state-add-scope-button.png" alt-text="Set state to enabled and select the add scope button.":::

    The new scope you defined displays on the pane.

    :::image type="content" source="../images/azure-portal-scope-added-successfully.png" alt-text="The new scope displayed on the Expose an API pane.":::

    > [!NOTE]
    > The domain part of the **Scope name** displayed just below the text field should automatically match the **Application ID URI** set in the previous step, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Select **Add a client application**.

    :::image type="content" source="../images/azure-portal-add-a-client-application.png" alt-text="Select add a client application.":::

    The **Add a client application** pane appears.

1. In the **Client ID** enter `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`. This value pre-authorizes all Microsoft Office application endpoints. If you also want to pre-authorize Office when used inside of Microsoft Teams, add `1fec8e78-bce4-4aaf-ab1b-5451cc387264` (Microsoft Teams desktop and Teams mobile) and `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams on the web).

    > [!NOTE]
    > The `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID pre-authorizes Office on all the following platforms. Alternatively, you can enter a proper subset of the following IDs if, for any reason, you want to deny authorization to Office on some platforms. If you do so, leave out the IDs of the platforms from which you want to withhold authorization. Users of your add-in on those platforms will not be able to call your Web APIs, but other functionality in your add-in will still work.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)



1. In **Authorized scopes**, select the `api://<fully-qualified-domain-name>/<app-id>/access_as_user` checkbox.

1. Select **Add application**.

    :::image type="content" source="../images/azure-portal-add-application.png" alt-text="The Add a client application pane.":::

## Add Microsoft Graph permissions

1. From the left pane, select **API permissions**.

    :::image type="content" source="../images/azure-portal-api-permissions.png" alt-text="The API permissions pane.":::

    The **API permissions** pane opens.

1. Select **Add a permission**.

    :::image type="content" source="../images/azure-portal-add-a-permission.png" alt-text="Adding a permission on the API permissions pane.":::

    The **Request API permissions** pane opens.

1. Select **Microsoft Graph**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-graph.png" alt-text="The Request API permissions pane with Microsoft Graph button.":::

1. Select **Delegated permissions**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-delegated.png" alt-text="The Request API permissions pane with delegated permissions button.":::

1. In the **Select permissions** search box, search for the permissions your add-in needs. For example, for an Outlook add-in, you might use `profile`, `openid`, `Files.ReadWrite`, and `Mail.Read`.

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It's a good practice to only request permissions that are needed, so we recommend that you uncheck the box for this permission if your add-in doesn't actually need it.

1. Select the checkbox for each permission as it appears. Note that the permissions will not remain visible in the list as you select each one. After selecting the permissions that your add-in needs, select **Add permissions**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-add-permissions.png" alt-text="The Request API permissions pane with some permissions selected.":::

1. Select **Grant admin consent for [tenant name]**. Select **Yes** for the confirmation that appears.

## Configure access token version

You must define the access token version that is acceptable for your app. This configuration is made in the Azure Active Directory application manifest.

### Define the access token version

The access token version can change if you chose an account type other than **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**. Use the following steps to ensure the access token version is correct for Office SSO usage.

1. From the left pane, select **Manifest**.

    :::image type="content" source="../images/azure-portal-manifest.png" alt-text="Select Azure manifest.":::

    The Azure Active Directory application manifest appears.

1. Enter **2** as the value for the `requestedAccessTokenVersion` property (in the `api` object).

    :::image type="content" source="../images/azure-portal-manifest-token-version.png" alt-text="Value for accepted access token version.":::

1. Select **Save**.

    A message pops up on the browser stating that the manifest was updated successfully.

    :::image type="content" source="../images/azure-portal-manifest-updated-message.png" alt-text="Manifest updated message.":::

Congratulations! You've completed the app registration to enable SSO for your Office add-in.
