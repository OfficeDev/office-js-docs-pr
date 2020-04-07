---
title: Outlook add-in requirements
description: 'For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.'
ms.date: 10/09/2019
localization_priority: Priority
---

# Outlook add-in requirements

For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.

## Client requirements

- The client must be one of the supported hosts for Outlook add-ins. The following clients support add-ins:

   - Outlook 2013 or later on Windows
   - Outlook 2016 or later on Mac
   - Outlook on iOS
   - Outlook on Android
   - Outlook on the web for Exchange 2016 or later and Office 365
   - Outlook on the web for Exchange 2013
   - Outlook.com

- The client must be connected to an Exchange server or Office 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.

## Mail server requirements

If the user is connected to Office 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.

- The server must be Exchange 2013 or later.
- Exchange Web Services (EWS) must be enabled and must be exposed to the Internet. Many add-ins require EWS to function properly.
- The server must have a valid authentication certificate in order for the server to issue valid identity tokens. New installations of Exchange Server include a default authentication certificate. For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).
- To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.

## Add-in server requirements

Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.

## See also

- [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md)
- [Office Add-in host and platform availability (Outlook section)](../overview/office-add-in-availability.md#outlook)
- [Outlook JavaScript API requirement set support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
