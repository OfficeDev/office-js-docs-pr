
> [!NOTE]
> This procedure is only needed when you're developing the add-in. When your production add-in is deployed to AppSource or an add-in catalog, users will individually trust it or an admin will consent for the organization at installation.

Carry out this procedure *after* you have [registered the add-in](../develop/register-sso-add-in-aad-v2.md).

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to view your app registration.

1. Sign in with the ***admin*** credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select the app with display name **$ADD-IN-NAME$**.

1. On the **$ADD-IN-NAME$** page, select **API permissions** then, under the **Grant consent** section, choose the **Grant admin consent for [tenant name]** button. Select **Yes** for the confirmation that appears.

> [!NOTE]
> We recommend this procedure as a best practice if you are using a Developer O365 tenant. However, if you prefer, it is possible to sideload an SSO add-in under development and prompt the user with a consent form. For more information, see [Sideload on Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) and [Sideload on Office Online](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).