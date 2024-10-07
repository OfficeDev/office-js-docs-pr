
> [!NOTE]
> This procedure is only needed when you're developing the add-in. When your production add-in is deployed to AppSource or the Microsoft 365 admin center, users will individually trust it or an admin will consent for the organization at installation.

Carry out this procedure *after* you have [registered the add-in](../develop/register-sso-add-in-aad-v2.md). (If you have just completed that procedure and the **API permissions** tab of the **$ADD-IN-NAME$** page is open in your browser, you can choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears. Skip the rest of this procedure.)

1. Browse to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to view your app registration.

1. Sign in with the ***admin*** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select the app with display name **$ADD-IN-NAME$**.

1. On the **$ADD-IN-NAME$** page, select **API permissions** then, under the **Configured permissions** section, choose **Grant admin consent for [tenant name]**. Select **Yes** for the confirmation that appears.

> [!NOTE]
> We recommend this procedure as a best practice if you're using a [Microsoft 365 developer account](https://aka.ms/m365devprogram). However, if you prefer, it is possible to sideload an SSO add-in under development and prompt the user with a consent form. For more information, see [Sideload on desktop](../testing/test-debug-non-local-server.md) and [Sideload on Office on the web](../testing/sideload-office-add-ins-for-testing.md).
