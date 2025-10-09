---
title: Get and set metadata in an Outlook add-in
description: Manage custom data in your Outlook add-in by using roaming settings, custom properties, or session data.
ms.date: 02/11/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get and set add-in metadata for an Outlook add-in

Manage custom data in your Outlook add-in using roaming settings, custom properties, or session data. These options give access to custom data that's only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings isn't accessible by custom properties, and vice versa.

The following table provides an overview of the available options to manage custom data in Outlook add-ins.

| Custom data option | Minimum requirement set | Applies to | Description |
| ----- | ----- | ----- | ----- |
| Roaming settings | [1.1](/javascript/api/requirement-sets/outlook/requirement-set-1.1/outlook-requirement-set-1.1) | Mailbox | Manages custom data in a user's mailbox. The add-in that sets the custom data can access it from other supported devices where the user's mailbox is set up. Stored data is accessible in subsequent Outlook sessions. |
| Custom properties | [1.1](/javascript/api/requirement-sets/outlook/requirement-set-1.1/outlook-requirement-set-1.1) | Mail item | Manages custom data for a mail item in a user's mailbox. The add-in that sets the custom data can access it from the mail item on supported devices where the user's mailbox is set up. Stored data is accessible in subsequent Outlook sessions. |
| Session data | [1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11) | Mail item | Manages custom data for a mail item in the user's current Outlook session. The add-in that sets the custom data can only access it from the mail item while it's being composed. |

> [!NOTE]
> For information on requirement sets and their supported clients, see [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

To learn more about each custom data option, select the applicable tab.

# [Roaming settings](#tab/roaming-settings)

You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).

Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they'll be available the next time the user opens your add-in, on the same or any other supported device.

> [!IMPORTANT]
> While the Outlook add-in API limits access to these settings to only the add-in that created them, these settings shouldn't be considered secure storage. They can be accessed by other services, such as Microsoft Graph. They shouldn't be used to store sensitive information, such as user credentials or security tokens.

### Roaming settings format

The data in a **RoamingSettings** object is stored as a serialized JavaScript Object Notation (JSON) string.

The following is an example of the structure, assuming there are three defined roaming settings named `add-in_setting_name_0`,  `add-in_setting_name_1`, and  `add-in_setting_name_2`.

```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```

### Loading roaming settings

A mail add-in typically loads roaming settings in the [Office.onReady](/javascript/api/office#office-office-onready-function(1)) handler. The following JavaScript code example shows how to load existing roaming settings and get the values of two settings, **customerName** and **customerBalance**.

```javascript
let _mailbox;
let _settings;
let _customerName;
let _customerBalance;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    _customerName = _settings.get("customerName");
    _customerBalance = _settings.get("customerBalance");
  }
});
```

### Creating or assigning a roaming setting

Continuing with the earlier example, the following JavaScript function, `setAddInSetting`, shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings) method to set a setting named `cookie` with today's date. Then, it persists the data by using the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)) method to save all the roaming settings to the user's mailbox.

The `set` method creates the setting if the setting doesn't already exist, and assigns the setting to the specified value. The `saveAsync` method saves roaming settings asynchronously. This code sample passes a callback function, `saveMyAddInSettingsCallback`, to `saveAsync`. When the asynchronous call finishes, `saveMyAddInSettingsCallback` is called by using one parameter, *asyncResult*. This parameter is an [AsyncResult](/javascript/api/office/office.asyncresult) object that contains the result of and any details about the asynchronous call. You can use the optional *userContext* parameter to pass any state information from the asynchronous call to the callback function.

```javascript
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings to the mailbox, so that they'll be available in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback function after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### Removing a roaming setting

Still extending the earlier example, the following JavaScript function, `removeAddInSetting`, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) method to remove the `cookie` setting and save all the roaming settings to the mailbox.

```javascript
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox, so that they'll be available in the next Outlook session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```

### Try the code example in Script Lab

To learn how to create and manage a RoamingSettings object, get the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and try out the ["Use add-in settings" sample](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/refs/heads/main/samples/outlook/10-roaming-settings/roaming-settings.yaml). To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

# [Custom properties](#tab/custom-properties)

You can specify data specific to an item in the user's mailbox using the [CustomProperties](/javascript/api/outlook/office.customproperties) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in doesn't offer to create the appointment a second time.

Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, use [CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)).

These add-in-specific, item-specific custom properties can only be accessed by using the `CustomProperties` object. These properties are different from the custom, MAPI-based [UserProperties](/office/vba/api/Outlook.UserProperties) in the Outlook object model, and extended properties in Exchange Web Services (EWS). You can't directly access `CustomProperties` by using the Outlook object model, EWS, or Microsoft Graph. To learn how to access `CustomProperties` using Microsoft Graph or EWS, see the section [Get custom properties using Microsoft Graph or EWS](#get-custom-properties-using-microsoft-graph-or-ews).

> [!NOTE]
> Custom properties are only available to the add-in that created them and only through the mail item in which they were saved. Because of this, properties set while in compose mode aren't transmitted to recipients of the mail item. When a message or appointment with custom properties is sent, its properties can be accessed from the item in the **Sent Items** folder. To allow recipients to receive the custom data your add-in sets, consider using [InternetHeaders](internet-headers.md) instead.

### Using custom properties

Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) method. After you've created the property bag, you can use the [set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) and [get](/javascript/api/outlook/office.customproperties) methods to add and retrieve custom properties. You must use the [saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) method to save any changes that you make to the property bag.

 > [!NOTE]
 > When using custom properties in Outlook add-ins, keep in mind that:
 >
 > - Outlook on Mac doesn't cache custom properties. If the user's network goes down, add-ins in Outlook on Mac wouldn't be able to access their custom properties.
 > - In classic Outlook on Windows, custom properties saved while in compose mode only persist after the item being composed is closed or after `Office.context.mailbox.item.saveAsync` is called.

### Try the code example in Script Lab

To learn how to create and manage a CustomProperties object, get the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and try out the ["Work with item custom properties" sample](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/refs/heads/main/samples/outlook/15-item-custom-properties/load-set-get-save.yaml). To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

### Get custom properties using Microsoft Graph or EWS

To get **CustomProperties** using Microsoft Graph or EWS, you should first determine the name of its MAPI-based extended property. You can then get that property in the same way you would get any MAPI-based extended property.

The use of Microsoft Graph or EWS depends on whether an add-in is running in an [Exchange Online](#exchange-online) or [Exchange on-premises](#exchange-on-premises) environment.

#### Exchange Online

In Exchange Online environments, your add-in can construct a Microsoft Graph request against messages and events to get the ones that already have custom properties. In your request, you should include the **CustomProperties** MAPI-based property and its property set using the details provided in [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).

The following example shows how to get all events that have any custom properties set by your add-in. It also ensures that the response includes the value of the property, so you can apply further filtering logic.

> [!IMPORTANT]
> In the following example, replace `<app-guid>` with your add-in's ID.

```http
GET https://graph.microsoft.com/v1.0/me/events?$filter=singleValueExtendedProperties/Any
  (ep: ep/id eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/value ne null)
  &$expand=singleValueExtendedProperties($filter=id eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

For other examples that get single-value MAPI-based extended properties, see [Get singleValueLegacyExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).

> [!TIP]
> To learn how to obtain an access code to Microsoft Graph, see [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md).

#### Exchange on-premises

In Exchange on-premises environments, your mail add-in can get the `CustomProperties` MAPI-based extended property using the EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation. Access `GetItem` on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method. In the `GetItem` request, specify the `CustomProperties` MAPI-based property in its property set using the details provided in [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).

The following example shows how to get an item and its custom properties.

> [!IMPORTANT]
> In the following example, replace `<app-guid>` with your add-in's ID.

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

You can also get more custom properties if you specify them in the request string as other [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elements.

#### How custom properties are stored on an item

Custom properties set by an add-in aren't equivalent to normal MAPI-based properties. Add-in APIs serialize all your add-in's `CustomProperties` as a JSON payload and then save them in a single MAPI-based extended property whose name is `cecp-<app-guid>` (`<app-guid>` is your add-in's ID) and property set GUID is `{00020329-0000-0000-C000-000000000046}`. (For more information about this object, see [MS-OXCEXT 2.2.5 Mail App Custom Properties](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481).) You can then use Microsoft Grpah or EWS to get this MAPI-based property.

### Platform behavior in messages

The following table summarizes saved custom properties behavior in messages for various Outlook clients.

|Scenario|Outlook on the web and on new Windows client|classic Outlook on Windows|Outlook on Mac|
|---|---|---|---|
|New compose|null|null|null|
|Reply, reply all|null|null|null|
|Forward|null|Loads parent's properties|null|
|Sent item from new compose|null|null|null|
|Sent item from reply or reply all|null|null|null|
|Sent item from forward|null|Removes parent's properties if not saved|null|

To handle the situation in classic Outlook on Windows:

1. Check for existing properties on initializing your add-in, and keep them or clear them as needed.
1. When setting custom properties, include an additional property to indicate whether the custom properties were added in read mode. This will help you differentiate if the property was created in compose mode or inherited from the parent.
1. To check if the user is forwarding or replying to a message, you can use [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getcomposetypeasync-member(1)) (available from requirement set 1.10).

# [Session data](#tab/session-data)

If you only need to save and access data while a mail item is being composed, use the [SessionData](/javascript/api/outlook/office.sessiondata) API. Because data is only saved for the duration of the current compose session, data from a SessionData object can't be accessed from an item that's been saved as a draft. This behavior applies even if the same add-in is used.

Custom data is saved to the SessionData object as key-value pairs. For each mail item, the data in the SessionData object is limited to 50,000 characters per add-in. That is, if multiple add-ins set custom session data on a single mail item, each add-in can create a SessionData object that contains up to 50,000 characters.

### Try the code example in Script Lab

To learn how to create and manage a SessionData object, get the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and try out the ["Work with session data APIs (Compose)" sample](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/refs/heads/main/samples/outlook/90-other-item-apis/session-data-apis.yaml). To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

---

## See also

- [MAPI Property Overview](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md)
- [Get and set internet headers on a message in an Outlook add-in](internet-headers.md)
