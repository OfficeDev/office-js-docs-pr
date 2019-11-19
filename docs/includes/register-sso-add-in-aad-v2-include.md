

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.

1. Sign in with the ***admin*** credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    * Set **Name** to **$ADD-IN-NAME$**.
    * Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.
    * Leave **Redirect URI** empty.
    * Choose **Register**.

1. On the **$ADD-IN-NAME$** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**. You'll use both of them in later procedures.

    > [!NOTE]
    > This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**. *Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.

1. Select **Expose an API** under **Manage**. Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$". Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > You may get an inaccurate error at this point saying "The application ID URI must be a valid URI starting with HTTPS, API, URN, MS-APPX. It must not end in a slash." If the ID meets the stated conditions, ignore the error and save your change.

    > [!NOTE]
    > If you get an error saying that the domain is already owned but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step. (This error can also occur if you are not signed in with credentials of an admin in the Office 365 tenancy. See step 2. Sign out and sign in again with admin credentials and repeat the process from step 3.)

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope name**.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user. Suggestions:

    - **Admin consent title:** Office can act as the user.
    - **Admin consent description:** Enable Office to call the add-in's web APIs with the same rights as the current user.
    - **User consent title:** Office can act as you.
    - **Admin consent description:** Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope**.

    > [!NOTE]
    > The domain part of the **Scope name** displayed just below the text field should automatically match the **Application ID URI** set in the previous step, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized.
  
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office on the web)

    For each ID, take these steps:

      a. Select **Add a client application** button then, in the panel that opens, set the **Client ID** to the respective GUID and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

      b. Select **Add application**.

1. Select **Authentication** under **Manage**. In the **Redirect URIs** section, select **Web** in the **Type** dropdown then set the **Redirect URI** value to `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Select **Save** at the top of the form.

1. Select **API permissions** under **Manage** and select **Add a permission**. On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search for the permissions your add-in needs. The following are examples.

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. Select the check box for each permission as it appears (note that the permissions will not remain visible in the list as you select each one). After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.
