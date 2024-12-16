---
title: Turn legacy Exchange Online tokens on or off
description: Turn legacy Exchange Online tokens on or off
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: how-to
ms.date: 12/12/2024
---

# Turn legacy Exchange Online tokens on or off

Legacy Exchange Online tokens are deprecated and will begin being turned off across Microsoft 365 tenants in February 2025. If you are a developer migrating your Outlook add-in from legacy tokens to Entra ID tokens and nested app authentication, you'll need to test updates to your add-in. You can use the Exchange Online PowerShell cmdlets to turn legacy tokens on or off. Turn off legacy tokens in a test tenant to confirm that your updated Outlook add-in is working correctly.

For more information about deprecation of legacy Exchange Online tokens, see [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

## Connect to Exchange Online PowerShell

To run the commands you need to connect to Exchange Online PowerShell.

1. Open PowerShell.
1. Run the command `Import-Module ExchangeOnlineManagement`. For more information about this command, see [Exchange Online PowerShell](/powershell/exchange/exchange-online-powershell).
1. To be sure you are on the latest version of the module, run the command `Update-Module -Name ExchangeOnlineManagement`.
1. Run the command `Connect-ExchangeOnline`. Sign in with your Microsoft 365 administrator credentials.

## Turn off legacy Exchange Online tokens in a test tenant

The `Set-AuthenticationPolicy` command controls the issuance of legacy Exchange Online tokens. When issuance is turned off, add-ins can no longer request user identity tokens or callback tokens. Existing tokens already issued will continue to work until they expire. It can take up to 24 hours before all request from Outlook add-ins for legacy Exchange Online tokens are blocked.

To turn legacy tokens off, run the following command.

`Set-AuthenticationPolicy –BlockLegacyExchangeTokens -Identity "LegacyExchangeTokens"`

The command turns off legacy tokens for the entire tenant. If an Outlook add-in requests a legacy token, it won’t be issued a token.

## Turn on legacy Exchange Online tokens in a test tenant

To turn legacy tokens on, run the following command. It can take up to 24 hours before all requests from Outlook add-ins for legacy tokens are allowed.

`Set-AuthenticationPolicy –AllowLegacyExchangeTokens -Identity "LegacyExchangeTokens"`

You’ll only be able to turn tokens back on until June 2025 when all legacy tokens in all tenants will be forced off. For more information, see the [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

> [!NOTE]
> It might take up to 24 hours for the change to take effect across your entire organization. Legacy Exchange tokens issued to Outlook add-ins before token blocking was implemented in your organization will remain valid until they expire.

## See also

- [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ)
- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md)
- [Set-AuthenticationPolicy](/powershell/module/exchange/set-authenticationpolicy)
- [Remove-AuthenticationPolicy](/powershell/module/exchange/remove-authenticationpolicy)
- [Get-AuthenticationPolicy](/powershell/module/exchange/get-authenticationpolicy)
