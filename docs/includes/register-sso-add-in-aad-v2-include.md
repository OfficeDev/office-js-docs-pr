

1. Navigate to [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Sign-in with the ***admin*** credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com

1. Click **Add an app**.

1. When prompted, enter **$ADD-IN-NAME$** as the app name, and then press **Create application**.

1. When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.

    > [!NOTE]
    > This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.

1. In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.

1. In the **Platforms** section, click **Add Platform**.

1. In the dialog that opens, select **Web API**.

1. An **Application ID URI** has been generated of the form “api://$App ID GUID$”. Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step. (This error can also occur if you are not signed in with credentials of an admin in the Office 365 tenancy. See step 2. Sign out and sign in again with admin credentials and repeat the process from step 3.)

    > [!NOTE]
    > The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized. Each time you enter one, a new empty textbox appears. (Enter only the GUID.)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office web client)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office web client)

1. Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

1. Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.

1. In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.

1. In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs. The following are examples:

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. At the bottom of the dialog, click **OK**.

1. At the bottom of the registration page, click **Save**.
