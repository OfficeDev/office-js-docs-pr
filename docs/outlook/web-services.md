---
title: Use Exchange Web Services (EWS) from an Outlook add-in
description: Provides an example that shows how an Outlook add-in can request information from Exchange Web Services.
ms.date: 04/28/2020
localization_priority: Normal
---

# Call web services from an Outlook add-in

Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.

The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.


**Table 1. Ways to call web services from an Outlook add-in**

<br/>

|**Web service location**|**Way to call the web service**|
|:-----|:-----|
|The Exchange server that hosts the client mailbox|Use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.|
|The web server that provides the source location for the add-in UI|Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.|
|All other locations|Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).|

## Using the makeEwsRequestAsync method to access EWS operations

You can use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to make an EWS request to the Exchange server that hosts the user's mailbox.

EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.

To use the `makeEwsRequestAsync` method to initiate an EWS operation, provide the following:

- The XML for the SOAP request for that EWS operation, as an argument to the  _data_ parameter

- A callback method (as the  _callback_ argument)

- Any optional input data for that callback method (as the  _userContext_ argument)

When the EWS SOAP request is complete, Outlook calls the callback method with one argument, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object. The callback method can access two properties of the `AsyncResult` object: the `value` property, which contains the XML SOAP response of the EWS operation, and optionally, the `asyncContext` property, which contains any data passed as the `userContext` parameter. Typically, the callback method then parses the XML in the SOAP response to get any relevant information, and processes that information accordingly.


## Tips for parsing EWS responses

When parsing a SOAP response from an EWS operation, note the following browser-dependent issues.


- Specify the prefix for a tag name when using the DOM method `getElementsByTagName`, to include support for Internet Explorer.

  `getElementsByTagName` behaves differently depending on browser type. For example, an EWS response can contain the following XML (formatted and abbreviated for display purposes).

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   Code, as in the following, would work on a browser like Chrome to get the XML enclosed by the `ExtendedProperty` tags.

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   On Internet Explorer, you must include the `t:` prefix of the tag name, as follows.

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- Use the DOM property `textContent` to get the contents of a tag in an EWS response, as follows.

   ```js
      content = $.parseJSON(value.textContent);
   ```

   Other properties such as `innerHTML` may not work on Internet Explorer for some tags in an EWS response.


## Example

The following example calls `makeEwsRequestAsync` to use the [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to get the subject of an item. This example includes the following three functions.

- `getSubjectRequest` &ndash; Takes an item ID as input, and returns the XML for the SOAP request to call `GetItem` for the specified item.

- `sendRequest` &ndash; Calls  `getSubjectRequest` to get the SOAP request for the selected item, then passes the SOAP request and the callback method, `callback`, to `makeEwsRequestAsync` to get the subject of the specified item.

- `callback` &ndash; Processes the SOAP response which includes any subject and other information about the specified item.


```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}
```


## EWS operations that add-ins support

Outlook add-ins can access a subset of operations that are available in EWS via the `makeEwsRequestAsync` method. If you are unfamiliar with EWS operations and how to use the `makeEwsRequestAsync` method to access an operation, start with a SOAP request example to customize your _data_ argument.

The following describes how you can use the `makeEwsRequestAsync` method.

1. In the XML, substitute any item IDs and relevant EWS operation attributes with appropriate values.

1. Include the SOAP request as an argument for the  _data_ parameter of `makeEwsRequestAsync`.

1. Specify a callback method and call `makeEwsRequestAsync`.

1. In the callback method, verify the results of the operation in the SOAP response.

1. Use the results of the EWS operation according to your needs.

The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Table 2. Supported EWS operations**

<br/>

|**EWS operation**|**Description**|
|:-----|:-----|
|[CopyItem operation](/exchange/client-developer/web-service-reference/copyitem-operation)|Copies the specified items and puts the new items in a designated folder in the Exchange store.|
|[CreateFolder operation](/exchange/client-developer/web-service-reference/createfolder-operation)|Creates folders in the specified location in the Exchange store.|
|[CreateItem operation](/exchange/client-developer/web-service-reference/createitem-operation)|Creates the specified items in the Exchange store.|
|[ExpandDL operation](/exchange/client-developer/web-service-reference/expanddl-operation)|Displays the full membership of distribution lists.|
|[FindConversation operation](/exchange/client-developer/web-service-reference/findconversation-operation)|Enumerates a list of conversations in the specified folder in the Exchange store.|
|[FindFolder operation](/exchange/client-developer/web-service-reference/findfolder-operation)|Finds subfolders of an identified folder and returns a set of properties that describe the set of subfolders.|
|[FindItem operation](/exchange/client-developer/web-service-reference/finditem-operation)|Identifies items that are located in a specified folder in the Exchange store.|
|[GetConversationItems operation](/exchange/client-developer/web-service-reference/getconversationitems-operation)|Gets one or more sets of items that are organized in nodes in a conversation.|
|[GetFolder operation](/exchange/client-developer/web-service-reference/getfolder-operation)|Gets the specified properties and contents of folders from the Exchange store.|
|[GetItem operation](/exchange/client-developer/web-service-reference/getitem-operation)|Gets the specified properties and contents of items from the Exchange store.|
|[GetUserAvailability operation](/exchange/client-developer/web-service-reference/getuseravailability-operation)|Provides detailed information about the availability of a set of users, rooms, and resources within a specified time period.|
|[MarkAsJunk operation](/exchange/client-developer/web-service-reference/markasjunk-operation)|Moves email messages to the Junk Email folder, and adds or removes senders of the messages from the blocked senders list accordingly.|
|[MoveItem operation](/exchange/client-developer/web-service-reference/moveitem-operation)|Moves items to a single destination folder in the Exchange store.|
|[ResolveNames operation](/exchange/client-developer/web-service-reference/resolvenames-operation)|Resolves ambiguous email addresses and display names.|
|[SendItem operation](/exchange/client-developer/web-service-reference/senditem-operation)|Sends email messages that are located in the Exchange store.|
|[UpdateFolder operation](/exchange/client-developer/web-service-reference/updatefolder-operation)|Modifies the properties of existing folders in the Exchange store.|
|[UpdateItem operation](/exchange/client-developer/web-service-reference/updateitem-operation)|Modifies the properties of existing items in the Exchange store.|

 > [!NOTE]
 > FAI (Folder Associated Information) items cannot be updated (or created) from an add-in. These hidden messages are stored in a folder and are used to store a variety of settings and auxiliary data.  Attempting to use the UpdateItem operation will throw an ErrorAccessDenied error: "Office extension is not allowed to update this type of item". As an alternative, you may use the [EWS Managed API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) to update these items from a Windows client or a server application. Caution is recommended as internal, service-type data structures are subject to change and could break your solution.


## Authentication and permission considerations for makeEwsRequestAsync

When you use the `makeEwsRequestAsync` method, the request is authenticated by using the email account credentials of the current user. The `makeEwsRequestAsync` method manages the credentials for you so that you do not have to provide authentication credentials with your request.

> [!NOTE]
> The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) cmdlet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the `makeEwsRequestAsync` method to make EWS requests.

Your add-in must specify the `ReadWriteMailbox` permission in its add-in manifest to use the `makeEwsRequestAsync` method. For information about using the `ReadWriteMailbox` permission, see the section [ReadWriteMailbox permission](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) in [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).

## See also

- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
- [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md)
- [EWS reference for Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Mail apps for Outlook and EWS in Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

See the following for creating backend services for add-ins using ASP.NET Web API.

- [Create a web service for an Office Add-in using the ASP.NET Web API](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [The basics of building an HTTP service using ASP.NET Web API](https://dotnet.microsoft.com/apps/aspnet/apis)