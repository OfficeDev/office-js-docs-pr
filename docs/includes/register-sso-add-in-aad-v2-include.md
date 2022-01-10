### Create the app registration

First, complete the steps in [Quickstart: Register an application with the Microsoft identity platform](/azure/active-directory/develop/quickstart-register-app) to create an initial app registration. After you complete the step [Add credentials](/azure/active-directory/develop/quickstart-register-app#add-credentials) return to this article and continue following the steps in [Expose a web API](#expose-a-web-api).

### Expose a web API

1. Be sure you are viewing the app registration you just created.
1. Select **Expose an API** under **Manage**. Select the **Set** link. This opens a **Set the App ID URI** box with a generated Application ID URI in the form `api://<application-id>`. Insert your fully qualified domain name before the `<application-id>`. The entire ID should have the form `api://<fully-qualified-domain-name>/<application-id>`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > If you get an error saying that the domain is already owned but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step. (This error can also occur if you are not signed in with credentials of an admin in the Microsoft 365 tenancy. See step 2. Sign out and sign in again with admin credentials and repeat the process from step 3.)

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope name**.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user. Suggestions:

    - **Admin consent display name:** Office can act as the user.
    - **Admin consent description:** Enable Office to call the add-in's web APIs with the same rights as the current user.
    - **User consent display name:** Office can act as you.
    - **User consent description:** Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope**.

    > [!NOTE]
    > The domain part of the **Scope name** displayed just below the text field should automatically match the **Application ID URI** set in the previous step, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized.
  
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
    * `08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

    > [!NOTE]
    > The ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` includes all of the other IDs listed and can be used singularly to pre-authorize all of the Office host endpoints for use with your service in the Office add-in SSO flow. 

    For each ID, take these steps.

      a. Select **Add a client application**. In the panel that opens, set the **Client ID** to the respective GUID and check the box for `api://<fully-qualified-domain-name>/<application-id>/access_as_user`.

      b. Select **Add application**.

1. Under **Manage** select **Authentication**, then choose **Add a platform**.

1. In the **Configure platforms** pane, select **Web**, and then set the **Redirect URI** value to `https://<fully-qualified-domain-name>`.

1. Choose **Configure**.

1. Under **Manage**, select **API permissions**, and select **Add a permission**. On the panel that opens, choose **Microsoft Graph**, and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search for the permissions your add-in needs. The following are examples.

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It is a good practice to only request permissions that are needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. Select the check box for each permission as it appears (note that the permissions will not remain visible in the list as you select each one). After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.
