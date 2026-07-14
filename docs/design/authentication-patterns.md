---
title: Authentication design guidelines for Office Add-ins
ms.date: 07/10/2026
ms.topic: best-practice
description: Learn design patterns for sign-in and sign-up experiences in Office Add-ins.
ai-usage: ai-assisted

ms.localizationpriority: medium
---

# Authentication patterns

Add-ins may require users to sign in or sign up to access features. Common authentication controls include username and password fields and buttons that start non-Microsoft credential flows. A simple, efficient authentication experience helps users get started quickly.

## Best practices

|Do|Don't|
|:----|:----|
|Before sign-in, describe the value of your add-in or demonstrate functionality without requiring an account. |Expect users to sign in without understanding the value and benefits of your add-in.|
|Guide users through authentication flows with a primary, highly visible button on each screen. |Draw attention to secondary and tertiary tasks with competing buttons and calls to action.|
|Use clear button labels that describe specific tasks like "Sign in" or "Create account". |Use vague button labels like "Submit" or "Get started".|
|Use a dialog to focus users' attention on authentication forms. |Overcrowd your task pane with a first-run experience and authentication forms.|
|Find small efficiencies in the flow, like auto-focusing input boxes. |Add unnecessary steps to the interaction, like requiring users to click into form fields.|
|Provide a way for users to sign out and reauthenticate. |Force users to uninstall to switch identities.|

## Authentication flow

Use this flow to guide users from first-run experience to a completed sign-in.

1. First-run placemat: Place your sign-in button as a clear call to action in your add-in's first-run experience.

    :::image type="content" source="../images/add-in-fre-value-placemat.png" alt-text="A sample add-in task pane in an Office application.":::

1. Identity provider choices dialog: Display a clear list of identity providers, including a username and password form if applicable. Your add-in UI might be blocked while the authentication dialog is open.

    :::image type="content" source="../images/add-in-auth-choices-dialog.png" alt-text="The Identity Provider Choices dialog in an Office application.":::

1. Identity provider sign-in: The identity provider has its own UI. Microsoft Entra ID supports customization of sign-in and access panel pages to keep the look and feel consistent with your service. For more information, see [Configure your company branding](/entra/fundamentals/how-to-customize-branding).

    :::image type="content" source="../images/add-in-auth-identity-sign-in.png" alt-text="The Identity Provider Sign-in dialog in an Office application.":::

1. Progress: Indicate progress while settings and UI load.

    :::image type="content" source="../images/add-in-auth-modal-interstitial.png" alt-text="A sample dialog with a progress indicator in an Office application.":::

> [!NOTE]
> When using Microsoft Entra ID, you can use a branded sign-in button that's customizable for light and dark themes. For more information, see [Microsoft identity platform and OAuth 2.0 authorization code flow](/entra/identity-platform/v2-oauth2-auth-code-flow).

## Single sign-on authentication flow

> [!NOTE]
> The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about single sign-on support, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets). If you're working with an Outlook add-in, be sure to enable modern authentication for the Microsoft 365 tenancy. For information about how to do this, see [Enable or disable modern authentication for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online).

Use single sign-on for a smoother end-user experience. The user's identity within Office (either a Microsoft Account or a Microsoft 365 identity) is used to sign in to your add-in. As a result users only sign in once. This removes friction in the experience making it easier for your customers to get started.

1. When an add-in is installed, the user sees a consent window similar to the following:

    :::image type="content" source="../images/add-in-auth-SSO-consent-dialog.png" alt-text="The consent window in an Office application when an add-in is being installed.":::

    > [!NOTE]
    > The add-in publisher controls the logo, strings, and permission scopes included in the consent window. Microsoft preconfigures the consent window UI.

1. After the user consents, the add-in loads and can extract and display user-specific information.

    :::image type="content" source="../images/add-in-ribbon.png" alt-text="An Office application with add-in buttons displayed on the ribbon.":::

## See also

- [Develop an Office Add-in with SSO](../develop/sso-in-office-add-ins.md)
