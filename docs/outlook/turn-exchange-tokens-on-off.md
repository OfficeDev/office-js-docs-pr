---
title: Turn legacy Exchange Online tokens on or off
description: Turn legacy Exchange Online tokens on or off
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: how-to
ms.date: 04/17/2025
---

# Turn legacy Exchange Online tokens on or off

Legacy Exchange Online tokens are deprecated and will be turned off across Microsoft 365 tenants starting February 17th, 2025. If you're a developer migrating your Outlook add-in from legacy tokens to Entra ID tokens and nested app authentication, you'll need to test updates to your add-in. Use the Exchange Online PowerShell cmdlets to turn off legacy tokens in a test tenant to confirm that your updated Outlook add-in is working correctly.

For more information about deprecation of legacy Exchange Online tokens, see [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

## Connect to Exchange Online PowerShell

To run the commands you need to connect to Exchange Online PowerShell.

1. Open PowerShell.
1. Run the command `Import-Module -Name ExchangeOnlineManagement`. For more information about this command, see [Exchange Online PowerShell](/powershell/exchange/exchange-online-powershell).
1. To be sure you are on the latest version of the module, run the command `Update-Module -Name ExchangeOnlineManagement`.
1. Run the command `Connect-ExchangeOnline`. Sign in with your Microsoft 365 administrator credentials.

## Turn off legacy Exchange Online tokens

The `Set-AuthenticationPolicy` command controls the issuance of legacy Exchange Online tokens. When issuance is turned off, add-ins can no longer request user identity tokens or callback tokens. Existing tokens already issued will continue to work until they expire. It can take up to 24 hours before all request from Outlook add-ins for legacy Exchange Online tokens are blocked.

To turn legacy tokens off, run the following command.

`Set-AuthenticationPolicy -BlockLegacyExchangeTokens -Identity "LegacyExchangeTokens"`

The command turns off legacy tokens for the entire tenant. If an Outlook add-in requests a legacy token, it won’t be issued a token.

> [!NOTE]
> If you've confirmed that your tenant is not using any add-ins that require legacy Exchange Online tokens, we recommend you turn off legacy Exchange Online tokens as a security best practice. For more information on how to determine if you tenant has add-ins using legacy tokens, see the [Nested app authentication and Outlook legacy tokens deprecation FAQ](faq-nested-app-auth-outlook-legacy-tokens.md).

## Turn on legacy Exchange Online tokens

To turn legacy tokens on, run the following command. It can take up to 24 hours before all requests from Outlook add-ins for legacy tokens are allowed.

`Set-AuthenticationPolicy -AllowLegacyExchangeTokens -Identity "LegacyExchangeTokens"`

Important notes about this command.

- Legacy Exchange tokens issued to Outlook add-ins before token blocking was implemented in your organization will remain valid until they expire.
- If you turn on legacy Exchange Online tokens, then they won't be turned off in February 2025 when Microsoft turns them off for all tenants. For more information, see [Nested app authentication and Outlook legacy tokens deprecation FAQ](faq-nested-app-auth-outlook-legacy-tokens.md).
- You’ll only be able to turn tokens back on until June 2025 when all legacy tokens in all tenants will be forced off. For more information, see the [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).
- Although the `-Identity` parameter is required, it doesn't affect any specific authentication policy. The command always applies to the entire organization regardless of what value you use. We show the value as `LegacyExchangeTokens` in the examples to keep the intent clear.

## Get the status of legacy Exchange Online tokens and add-ins that use them

To view the status of legacy Exchange Online tokens, run the following command.

`Get-AuthenticationPolicy -AllowLegacyExchangeTokens`

The command returns whether `AllowLegacyExchangeTokens` is true or false, such as the following example in PowerShell.

```console
PS C:\> Get-AuthenticationPolicy -AllowLegacyExchangeTokens
AllowLegacyExchangeTokens: False
Allowed: []
Blocked: []
PS C:\>
```

> [!NOTE]
> The **Allowed** and **Blocked** lists aren't used. If you see any IDs listed, the IDs are old data that is no longer applicable.

> [!NOTE]
> The `Get-AuthenticationPolicy -AllowLegacyExchangeTokens` command is the only way to view legacy token status. Other commands, such as `Get-AuthenticationPolicy | Format-Table -Auto Name`, don't return the legacy token status.

The `Get-AuthenticationPolicy` command only shows the legacy token status as set by the administrator. If the administrator has never changed the settings, the command returns `(Not Set)`. If the token status is `(Not Set)` when the February deployment by Microsoft to turn off legacy tokens is implemented, the token status will still be `(Not Set)` even though legacy tokens are off. The following table shows the behavior of legacy Exchange Online tokens based on the token status when the change is applied.

| Legacy token admin setting  | Legacy token behavior before February change  | Legacy token behavior after February change | Legacy token behavior after June change |
|----------|------------|-------------|------------|
|(Not Set) | Tokens on  | Tokens off  | Tokens off |
|False     | Tokens off | Tokens off  | Tokens off |
|True      | Tokens on  | Tokens on   | Tokens off |

## See also

- [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ)
- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md)
- [Set-AuthenticationPolicy](/powershell/module/exchange/set-authenticationpolicy)
- [Remove-AuthenticationPolicy](/powershell/module/exchange/remove-authenticationpolicy)
- [Get-AuthenticationPolicy](/powershell/module/exchange/get-authenticationpolicy)
