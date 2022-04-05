## Create an app registration

Registering your application (the add-in) establishes a trust relationship between your add-in and the Microsoft identity platform. The trust is unidirectional: your add-in trusts the Microsoft identity platform, and not the other way around.

1. Sign in to the [Azure portal](https://portal.azure.com/) with the ***admin*** credentials to your Microsoft 365 tenancy. For example, **MyName@contoso.onmicrosoft.com**.
1. Under **Manage**, select **App registrations** > **New registration**. On the **Register an application** page, set the values as follows.

    * Set **Name** to `<add-in-name>`.
    * Set **Supported account types** to **Accounts in any organizational directory (any Azure AD directory - multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    * Leave **Redirect URI** empty.
    * Choose **Register**.

1. Copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**. You'll use both of them in later procedures.

    > [!NOTE]
    > This ID is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.

## Add a client secret

Sometimes called an _application password_, a client secret is a string value your app can use in place of a certificate to identity itself.

1. In the Azure portal, in **App registrations**, select your application.
1. Select **Certificates & secrets** > **Client secrets** > **New client secret**.
1. Add a description for your client secret.
1. Select an expiration for the secret or specify a custom lifetime.
    * Client secret lifetime is limited to two years (24 months) or less. You can't specify a custom lifetime longer than 24 months.
    * Microsoft recommends that you set an expiration value of less than 12 months.
1. Select **Add**.
1. _Record the secret's value_ for use in your client application code. This secret value is _never displayed again_ after you leave this page.

## Expose a web API

1. Be sure you are viewing the app registration you just created.
1. Under **Manage**, select **Expose an API**, and select the **Set** link. This opens a **Set the App ID URI** box with a generated Application ID URI in the form `api://<application-id>`. Insert your fully qualified domain name before the `<application-id>`. The entire ID should have the form `api://<fully-qualified-domain-name>/<application-id>`; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > If you get an error saying that the domain is already owned but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step. (This error can also occur if you are not signed in with credentials of an admin in the Microsoft 365 tenancy. See step 2. Sign out and sign in again with admin credentials and repeat the process from step 3.)

## Add a scope

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope name**.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user. Suggestions:

    * **Admin consent display name:** Office can act as the user.
    * **Admin consent description:** Enable Office to call the add-in's web APIs with the same rights as the current user.
    * **User consent display name:** Office can act as you.
    * **User consent description:** Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope**.

    > [!NOTE]
    > The domain part of the **Scope name** displayed just below the text field should automatically match the **Application ID URI** set in the previous step, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Authorized client applications** section, enter the following ID to pre-authorize all Microsoft Office application endpoints.

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (All Microsoft Office application endpoints)

    > [!NOTE]
    > The `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID pre-authorizes Office on all the following platforms. Alternatively, you can enter a proper subset of the following IDs if for any reason you want to deny authorization to Office on some platforms. Just leave out the IDs of the platforms from which you want to withhold authorization. Users of your add-in on those platforms will not be able to call your Web APIs, but other functionality in your add-in will still work.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. Select **Add a client application**. In the panel that opens, set the **Client ID** to the respective GUID and check the box for `api://<fully-qualified-domain-name>/<application-id>/access_as_user`.

1. Select **Add application**.

## Add Microsoft Graph permissions

1. Under **Manage**, select **Authentication**, then choose **Add a platform**.

1. In the **Configure platforms** pane, select **Web**, and then set the **Redirect URI** value to `https://<fully-qualified-domain-name>`.

1. Choose **Configure**.

1. Under **Manage**, select **API permissions**, and select **Add a permission**. On the panel that opens, choose **Microsoft Graph**, and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search for the permissions your add-in needs. The following are examples.

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It's a good practice to only request permissions that are needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. Select the check box for each permission as it appears (note that the permissions will not remain visible in the list as you select each one). After selecting the permissions that your add-in needs, select the **Add permissions** button.
