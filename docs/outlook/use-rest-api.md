---
title: Use the Outlook REST APIs from an Outlook add-in
description: Learn how to use the Outlook REST APIs from an Outlook add-in to get an access token.
ms.date: 09/02/2022
ms.localizationpriority: medium
---

# Use the Outlook REST APIs from an Outlook add-in

The [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.

> [!IMPORTANT]
> **The Outlook REST APIs are deprecated**
>
> The Outlook REST endpoints will be fully decommissioned on November 30, 2022 (for more details, see the [November 2020 announcement](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)). You should migrate existing add-ins to use [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph). For guidance, see [Compare Microsoft Graph and Outlook REST API endpoints](/outlook/rest/compare-graph).
>
> To assist you with the migration, active add-ins that use the REST service are eligible for an exemption to keep using the service until [extended support ends for Outlook 2019 on October 14, 2025](/lifecycle/end-of-support/end-of-support-2025). This includes new add-ins developed after November 30, 2022. The exemption is based on the add-in's manifest ID and applies to privately released and AppSource-hosted add-ins.
>
> Automatic traffic identification of Outlook add-ins that use the REST service is currently being tested for exemption validation. If you'd like to participate in this testing phase, please complete the [REST API add-in verification form](https://aka.ms/RESTCheck) prior to November 2022. For more information, see the [Office Add-ins August 2022 community call blog post](https://pnp.github.io/blog/office-add-ins-community-call/2022-08-10/).

## Get an access token

The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method introduced in the Mailbox requirement set 1.5.

By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.

### Add-in permissions and token scope

It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.

If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.

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

To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST. This is obtained from the [Office.context.mailbox.item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.

- In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.
- In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method.
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

The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) property.

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

- For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.
- Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token. For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).
