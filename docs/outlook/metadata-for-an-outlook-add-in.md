---
title: Get and set metadata in an Outlook add-in
description: Manage custom data in your Outlook add-in by using either roaming settings or custom properties.
ms.date: 07/08/2022
ms.localizationpriority: medium
---

# Get and set add-in metadata for an Outlook add-in

You can manage custom data in your Outlook add-in by using either of the following:

- Roaming settings, which manage custom data for a user's mailbox.
- Custom properties, which manage custom data for an item in a user's mailbox.

Both of these give access to custom data that is only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings is not accessible by custom properties, and vice versa. The data is stored on the server for that mailbox, and is accessible in subsequent Outlook sessions on all the form factors that the add-in supports.

## Custom data per mailbox: roaming settings

You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).

Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they will be available the next time the user opens your add-in, on the same or any other supported device.

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

A mail add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office#Office_initialize_reason_) event handler. The following JavaScript code example shows how to load existing roaming settings and get the values of two settings, **customerName** and **customerBalance**.

```js
let _mailbox;
let _settings;
let _customerName;
let _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}
```

### Creating or assigning a roaming setting

Continuing with the preceding example, the following JavaScript function,  `setAddInSetting`, shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings) method to set a setting named `cookie` with today's date, and persist the data by using the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)) method to save all the roaming settings back to the server.

The `set` method creates the setting if the setting does not already exist, and assigns the setting to the specified value. The `saveAsync` method saves roaming settings asynchronously. This code sample passes a callback function, `saveMyAddInSettingsCallback`, to `saveAsync` When the asynchronous call finishes,  `saveMyAddInSettingsCallback` is called by using one parameter, _asyncResult_. This parameter is an [AsyncResult](/javascript/api/office/office.asyncresult) object that contains the result of and any details about the asynchronous call. You can use the optional _userContext_ parameter to pass any state information from the asynchronous call to the callback function.

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback function after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### Removing a roaming setting

Also extending the preceding examples, the following JavaScript function,  `removeAddInSetting`, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.

```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```

## Custom data per item in a mailbox: custom properties

You can specify data specific to an item in the user's mailbox using the [CustomProperties](/javascript/api/outlook/office.customproperties) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in doesn't offer to create the appointment a second time.

Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, use [CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)).

These add-in-specific, item-specific custom properties can only be accessed by using the `CustomProperties` object. These properties are different from the custom, MAPI-based [UserProperties](/office/vba/api/Outlook.UserProperties) in the Outlook object model, and extended properties in Exchange Web Services (EWS). You cannot directly access `CustomProperties` by using the Outlook object model, EWS, or REST. To learn how to access `CustomProperties` using EWS or REST, see the section [Get custom properties using EWS or REST](#get-custom-properties-using-ews-or-rest).

### Using custom properties

Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) method. After you have created the property bag, you can use the [set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) and [get](/javascript/api/outlook/office.customproperties) methods to add and retrieve custom properties. You must use the [saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) method to save any changes that you make to the property bag.

 > [!NOTE]
 > Because Outlook on Mac doesn't cache custom properties, if the user's network goes down, mail add-ins in Outlook on Mac would not be able to access their custom properties.

### Custom properties example

The following example shows a simplified set of functions and methods for an Outlook add-in that uses custom properties. You can use this example as a starting point for your add-in that uses custom properties.

This example includes the following functions and methods.

- [Office.initialize](/javascript/api/office#Office_initialize_reason_) -- Initializes the add-in and loads the custom property bag from the Exchange server.

- **customPropsCallback** -- Gets the custom property bag that is returned from the server and saves it for later use.

- **updateProperty** -- Sets or updates a specific property, and then saves the change to the server.

- **removeProperty** -- Removes a specific property from the property bag, and then saves the removal to the server.

```js
let _mailbox;
let _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  const myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### Get custom properties using EWS or REST

To get **CustomProperties** using EWS or REST, you should first determine the name of its MAPI-based extended property. You can then get that property in the same way you would get any MAPI-based extended property.

#### How custom properties are stored on an item

Custom properties set by an add-in are not equivalent to normal MAPI-based properties. Add-in APIs serialize all your add-in's `CustomProperties` as a JSON payload and then save them in a single MAPI-based extended property whose name is `cecp-<app-guid>` (`<app-guid>` is your add-in's ID) and property set GUID is `{00020329-0000-0000-C000-000000000046}`. (For more information about this object, see [MS-OXCEXT 2.2.5 Mail App Custom Properties](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481).) You can then use EWS or REST to get this MAPI-based property.

#### Get custom properties using EWS

Your mail add-in can get the `CustomProperties` MAPI-based extended property by using the EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation. Access `GetItem` on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method. In the `GetItem` request, specify the `CustomProperties` MAPI-based property in its property set using the details provided in the preceding section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).

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

#### Get custom properties using REST

In your add-in, you can construct your REST query against messages and events to get the ones that already have custom properties. In your query, you should include the **CustomProperties** MAPI-based property and its property set using the details provided in the section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).

The following example shows how to get all events that have any custom properties set by your add-in and ensure that the response includes the value of the property so you can apply further filtering logic.

> [!IMPORTANT]
> In the following example, replace `<app-guid>` with your add-in's ID.

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

For other examples that use REST to get single-value MAPI-based extended properties, see [Get singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).

The following example shows how to get an item and its custom properties. In the callback function for the `done` method, `item.SingleValueExtendedProperties` contains a list of the requested custom properties.

> [!IMPORTANT]
> In the following example, replace `<app-guid>` with your add-in's ID.

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## See also

- [MAPI Property Overview](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Outlook Properties Overview](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [Call Outlook REST APIs from an Outlook add-in](use-rest-api.md)
- [Call web services from an Outlook add-in](web-services.md)
- [Properties and extended properties in EWS in Exchange](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [Property sets and response shapes in EWS in Exchange](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)
