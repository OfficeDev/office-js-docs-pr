---
title: Use the Microsoft Graph REST API from an Outlook add-in
description: Learn how to use the Outlook mail REST API from an Outlook add-in with Microsoft Graph.
ms.date: 04/22/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Use the Microsoft Graph REST API from an Outlook add-in

The Outlook JavaScript API (Office.js) retrieves the properties of messages and appointments and run operations on these items in your add-in. However, there may be instances where the data you need isn't available through the API. For example, your add-in may need to implement single sign-on or identify messages in a user's mailbox that originated from the same sender. To get the information you need, use the [Outlook mail REST API](/graph/api/resources/mail-api-overview) through [Microsoft Graph](/graph/overview).

## Get started

To make calls to the Microsoft Graph API in Office Add-ins, implement the [nested app authentication (NAA)](../develop/enable-nested-app-authentication-in-your-add-in.md) solution to request a token.

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

## Call the Microsoft Graph API

Once you have an access token, you can then use it to call Microsoft Graph.

The Microsoft Graph API consists of the v1.0 and beta endpoints. Note the following about the endpoint pattern.

- `version` specifies the `v1.0` or `beta` API.
- `resource` specifies the resource your add-in interacts with, such as a user, group, or site.
- `query_parameters` specifies parameters to customize your request. For example, you can filter the messages returned to only those from a specific sender.

```http
https://graph.microsoft.com/[version]/[resource]?[query_parameters]
```

For more information on the Microsoft Graph API and its components, see [Use the Microsoft Graph API](/graph/use-the-api).

## See also

- [Microsoft Graph REST API v1.0 endpoint reference](/graph/api/overview)
- [Outlook mail API overview](/graph/outlook-mail-concept-overview)
- [Use the Outlook mail REST API](/graph/api/resources/mail-api-overview)
- [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md)
