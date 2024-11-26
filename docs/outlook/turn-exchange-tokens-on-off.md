---
title: Turn legacy Exchange Online tokens on or off
description: Turn legacy Exchange Online tokens on or off
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: how-to
ms.date: 11/26/2024
---

# Turn legacy Exchange Online tokens on or off

Legacy Exchange Online tokens are deprecated and will begin being turned off across Microsoft 365 tenants in February 2025. If you are a developer migrating your Outlook add-in from legacy tokens to Entra ID tokens and nested app authentication, you'll need to test updates to your add-in. You can use the Exchange Online PowerShell cmdlets to turn legacy Exchange Online tokens on or off. Turn off legacy tokens in a test tenant to confirm that your updated Outlook add-in is working correctly.

For more information about deprecation of legacy Exchange Online tokens, see [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

> [!WARNING]
> The commands in this article are for testing purposes only. Don't use these commands on a production tenant. They can turn off some essential Outlook services that can cause breaking issues for users.

## Turn off legacy Exchange Online tokens in a test tenant

The `Set-AuthenticationPolicy` command controls the issuance of legacy Exchange Online tokens. When issuance is turned off, add-ins can no longer request user identity tokens or callback tokens. Existing tokens already issued will continue to work until they expire. It can take up to 24 hours before all request from Outlook add-ins for legacy Exchange Online tokens are blocked.

To turn legacy Exchange online tokens off, run the following command.

`Set-AuthenticationPolicy –BlockLegacyExchangeTokens -Identity "LegacyExchangeTokens"`

The command turns off legacy Exchange tokens for the entire tenant. If an Outlook add-in requests a legacy Exchange token, it won’t be issued a token.

## Turn on legacy Exchange Online tokens in a test tenant

To turn legacy Exchange online tokens on, run the following command. It can take up to 24 hours before all requests from Outlook add-ins for legacy Exchange Online tokens are allowed.

`Set-AuthenticationPolicy –AllowLegacyExchangeTokens -Identity "LegacyExchangeTokens"`

You’ll only be able to turn on tokens back on until June 2025 when all legacy Exchange online tokens in all tenants will be forced off. For more information, see the [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

## See also

- [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ)
- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md)
