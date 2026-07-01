---
title: Use the Microsoft Graph REST API from an Outlook add-in
description: Learn how to access Microsoft Graph data from your Outlook add-in, including how to get an access token using nested app authentication and call the Microsoft Graph API.
ms.date: 06/26/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Use the Microsoft Graph REST API from an Outlook add-in

The Outlook JavaScript API (Office.js) retrieves the properties of messages and appointments and runs operations on these items in your add-in. However, there may be instances where the data you need isn't available through the API. For example, your add-in may need to implement single sign-on or identify messages in a user's mailbox that originated from the same sender. To get the information you need, use the [Outlook mail REST API](/graph/api/resources/mail-api-overview) through [Microsoft Graph](/graph/overview).

## Get started

To call Microsoft Graph from your add-in, implement the [nested app authentication (NAA)](../develop/enable-nested-app-authentication-in-your-add-in.md) solution to get an access token. NAA uses the Microsoft Authentication Library (MSAL) to silently acquire a token for the signed-in user without requiring a redirect to a sign-in page. The token is scoped to the Microsoft Graph API endpoints your add-in needs.

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

## Call the Microsoft Graph API

Once you have an access token, use it to call Microsoft Graph by passing it as a Bearer token in the `Authorization` header.

The Microsoft Graph API uses the following endpoint pattern.

- `version` specifies `v1.0` (stable) or `beta` (preview).
- `resource` specifies the resource your add-in interacts with, such as a user, group, or site.
- `query_parameters` are optional parameters to filter, sort, or paginate results.

```http
https://graph.microsoft.com/[version]/[resource]?[query_parameters]
```

The following example shows how to get the signed-in user's messages using the Fetch API after acquiring an access token.

```javascript
// Call the Microsoft Graph API to get the user's messages.
const response = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
    headers: {
        'Authorization': 'Bearer ' + accessToken
    }
});
const data = await response.json();
console.log(data.value); // Array of message objects
```

For more information on the Microsoft Graph API and its components, see [Use the Microsoft Graph API](/graph/use-the-api).

## See also

- [Microsoft Graph REST API v1.0 endpoint reference](/graph/api/overview)
- [Outlook mail API overview](/graph/outlook-mail-concept-overview)
- [Use the Outlook mail REST API](/graph/api/resources/mail-api-overview)
- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md)
