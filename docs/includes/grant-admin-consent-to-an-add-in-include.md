
> [!NOTE]
> This procedure is only needed when you're developing the add-in. When your production add-in is deployed to AppSource or an add-in catalog, users will individually trust it or an admin will consent for organization at installation.

Carry out this procedure *after* you have [registered the add-in](../develop/register-sso-add-in-aad-v2.md).

1. In the following string, replace the placeholder “{application_ID}” with the Application ID that you copied when you registered your add-in:
    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Paste the resulting URL into a browser address bar and navigate to it.

1. When prompted, sign in with the admin credentials to your Office 365 tenancy.

1. You are then prompted to grant permission for your add-in to access your Microsoft Graph data. Click **Accept**.

1. The browser window/tab is then redirected to the **Redirect URL** that you specified when you registered the add-in. If the add-in's web application is running, the home page of the add-in opens in the browser; otherwise, you'll get a 404 error. But the fact that the browser attempted to open the home page means that consent was successfully granted.

>[!NOTE]
>We recommend this procedure as a best practice if you are using a Developer O365 tenant. However, if you prefer, it is possible to sideload an SSO add-in under development and prompt the user with a consent form. For more information, see [Sideload on Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) and [Sideload on Office](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
