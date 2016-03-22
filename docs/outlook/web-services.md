
# Call web services from an Outlook add-in

Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.

The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.


**Table 1. Ways to call web services from an Outlook add-in**


|**Web service location**|**Way to call the web service**|
|:-----|:-----|
|The Exchange server that hosts the client mailbox|Use the [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.|
|The web server that provides the source location for the add-in UI|Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.|
|All other locations|Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../../docs/develop/privacy-and-security.md).|

## Using the makeEwsRequestAsync method to access EWS operations


You can use the [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) method to make a EWS request to the Exchange server that hosts the user's mailbox.

EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform a EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS. 

To use the  **makeEwsRequestAsync** method to initiate a EWS operation, provide the following:


- The XML for the SOAP request for that EWS operation, as an argument to the  _data_ parameter
    
- A callback method (as the  _callback_ argument)
    
- Any optional input data for that callback method (as the  _userContext_ argument)
    
When the EWS SOAP request is complete, Outlook calls the callback method with one argument, which is an [AsyncResult](../../reference/outlook/simple-types.md) object. The callback method can access two properties of the **AsyncResult** object: the **value** property, which contains the XML SOAP response of the EWS operation, and optionally, the **asyncContext** property, which contains any data passed as the **userContext** parameter. Typically, the callback method then parses the XML in the SOAP response to get any relevant information, and processes that information accordingly.


## Tips for parsing EWS responses


When parsing a SOAP response from a EWS operation, note the following browser-dependent issues:


- Specify the prefix for a tag name when using the DOM method  **getElementsByTagName**, to include support for Internet Explorer.
    
     **getElementsByTagName** behaves differently depending on browser type. For example, a EWS response can contain the following XML (formatted and abbreviated for display purposes):
    
    ```XML
      <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
    PropertyName="MyProperty" 
    PropertyType="String"/>
    <t:Value>{
    ...
    }</t:Value></t:ExtendedProperty>
    ```

 Code as in the following would work on a browser like Chrome to get the XML enclosed by the  **ExtendedProperty** tags:

    ```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("ExtendedProperty");
    ```


   
 On Internet Explorer, you must include the  `t:` prefix of the tag name, as shown below:

    ```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
    ```

- Use the DOM property  **textContent** to get the contents of a tag in a EWS response, as shown below:
    
    ```
      content = $.parseJSON(value.textContent);
    ```

 Other properties such as  **innerHTML** may not work on Internet Explorer for some tags in a EWS response.
    

## Example


The following example calls  **makeEwsRequestAsync** to use the [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) operation to get the subject of an item. This example includes the following three functions:


-  `getSubjectRequest` -- Takes an item ID as input, and returns the XML for the SOAP request to call **GetItem** for the specified item.
    
-  `sendRequest` -- Calls  `getSubjectRequest` to get the SOAP request for the selected item, then passes the SOAP request and the callback method, `callback`, to  **makeEwsRequestAsync** to get the subject of the specified item.
    
-  `callback` -- Processes the SOAP response which includes any subject and other information about the specified item.
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
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


Outlook add-ins can access a subset of operations that are available in EWS via the  **makeEwsRequestAsync** method. If you are unfamiliar with EWS operations and how to use the **makeEwsRequestAsync** method to access an operation, start with a SOAP request example to customize your _data_ argument. The following describes how you can use the **makeEwsRequestAsync** method:


1. In the XML, substitute any item IDs and relevant EWS operation attributes with appropriate values.
    
2. Include the SOAP request as an argument for the  _data_ parameter of **makeEwsRequestAsync**.
    
3. Specify a callback method and call  **makeEwsRequestAsync**.
    
4. In the callback method, verify the results of the operation in the SOAP response.
    
5. Use the results of the EWS operation according to your needs.
    
The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx).


**Table 2. Supported EWS operations**


|**EWS operation**|**Description**|
|:-----|:-----|
|[CopyItem operation](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|Copies the specified items and puts the new items in a designated folder in the Exchange store.|
|[CreateFolder operation](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|Creates folders in the specified location in the Exchange store.|
|[CreateItem operation](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|Creates the specified items in the Exchange store.|
|[FindConversation operation](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|Enumerates a list of conversations in the specified folder in the Exchange store.|
|[FindFolder operation](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|Finds subfolders of an identified folder and returns a set of properties that describe the set of subfolders.|
|[FindItem operation](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|Identifies items that are located in a specified folder in the Exchange store.|
|[GetConversationItems operation](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|Gets one or more sets of items that are organized in nodes in a conversation.|
|[GetFolder operation](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|Gets the specified properties and contents of folders from the Exchange store.|
|[GetItem operation](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|Gets the specified properties and contents of items from the Exchange store.|
|[MarkAsJunk operation](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|Moves email messages to the Junk Email folder, and adds or removes senders of the messages from the blocked senders list accordingly.|
|[MoveItem operation](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|Moves items to a single destination folder in the Exchange store.|
|[SendItem operation](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|Sends email messages that are located in the Exchange store.|
|[UpdateFolder operation](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|Modifies the properties of existing folders in the Exchange store.|
|[UpdateItem operation](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|Modifies the properties of existing items in the Exchange store.|

## Authentication and permission considerations for the makeEwsRequestAsync method


When you use the  **makeEwsRequestAsync** method, the request is authenticated by using the email account credentials of the current user. The **makeEwsRequestAsync** method manages the credentials for you so that you do not have to provide authentication credentials with your request.


 >**Note**  The server administrator must use the [New-WebServicesVirtualDirctory](http://technet.microsoft.com/en-us/library/bb125176.aspx) or the [Set-WebServicesVirtualDirecory](http://technet.microsoft.com/en-us/library/aa997233.aspx) cmldet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the **makeEwsRequestAsync** method to make EWS requests.

Your add-in must specify the  **ReadWriteMailbox** permission in its add-in manifest to use the **makeEwsRequestAsync** method. For information about using the **ReadWriteMailbox** permission, see the section [ReadWriteMailbox permission](../outlook/understanding-outlook-add-in-permissions.md#olowa15conagave_permmodelreadwrite) in [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md).


## Additional resources



- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Privacy and security for Office Add-ins](../../docs/develop/privacy-and-security.md)
    
- [Addressing same-origin policy limitations in Office Add-ins](../../docs/develop/addressing-same-origin-policy-limitations.md)
    
- [EWS reference for Exchange](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- [Mail apps for Outlook and EWS in Exchange](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
See the following for creating backend services for add-ins using ASP.NET Web API:


- [Create a web service for an Office Add-in using the ASP.NET Web API](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [The basics of building an HTTP service using ASP.NET Web API](http://www.asp.net/web-api)
    
