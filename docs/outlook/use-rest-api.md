---
title: Use the Outlook REST APIs from an Outlook add-in
description: The Outlook REST API v2.0 endpoints are deprecated. Learn how to migrate to Microsoft Graph. For add-ins pending migration, this article covers getCallbackTokenAsync, item ID retrieval, and REST API calls.
ms.date: 04/22/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Use the Outlook REST APIs from an Outlook add-in

The Outlook REST API v2.0 endpoints are deprecated. For new add-in development, use the [Microsoft Graph REST API](microsoft-graph.md) instead. This article covers the REST API approach for existing add-ins that haven't yet migrated to Microsoft Graph.

## Get an access token

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getcallbacktokenasync-member(1)) method introduced in the Mailbox requirement set 1.5.

By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.

### Add-in permissions and token scope

It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the [read/write item permission](understanding-outlook-add-in-permissions.md#readwrite-item-permission) level in its manifest.

If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the [read/write mailbox permission](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission) level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.

### Example

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    const accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## Get the item ID

To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST. This is obtained from the `itemId` property ([MessageRead](/javascript/api/outlook/office.messageread#outlook-office-messageread-itemid-member), [AppointmentRead](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-itemid-member)), but some checks should be made to ensure that it is a REST-formatted ID.

- In Outlook on mobile devices, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.
- In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-converttorestid-member(1)) method.
- Note you must also convert Attachment ID to a REST-formatted ID in order to use it. The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.

Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostname-member) property.

### Example

```js
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## Get the REST API URL

The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-resturl-member) property.

### Example

```js
// Example: https://outlook.office.com
const restHost = Office.context.mailbox.restUrl;
```

## Call the API

After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX. The following example calls the Outlook Mail REST API to get the current message.

> [!IMPORTANT]
> For on-premises Exchange deployments, client-side requests using AJAX or similar libraries fail because CORS isn't supported in that server setup.

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  const itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  const getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    const subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## See also

- [Use the Microsoft Graph REST API from an Outlook add-in](microsoft-graph.md)
- [Authentication and authorization in Outlook add-ins](authentication.md)
- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md)
- [Add-in Command sample](https://github.com/OfficeDev/outlook-add-in-command-demo)
