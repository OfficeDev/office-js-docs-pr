---
title: Get an Outlook item's attachments from Exchange
description: Learn how your Outlook add-in can directly get attachments and their contents from Exchange Online or Exchange Server.
ms.date: 10/29/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Get an Outlook item's attachments from Exchange

The Office JavaScript API includes APIs to get attachments and their contents from messages and appointments in Outlook. The following table lists these attachment APIs, the Outlook modes they operate in, and the minimum Mailbox requirement set they need to operate.

| API | Supported Outlook modes | Minimum requirement set |
| ---- | ---- | ---- |
| [Office.context.mailbox.item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) | Read | [1.1](/javascript/api/requirement-sets/outlook/requirement-set-1.1/outlook-requirement-set-1.1) |
| [Office.context.mailbox.item.getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) | Compose | [1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) |
| [Office.context.mailbox.item.getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) | Read<br/>Compose | [1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) |

If the Outlook client in which the add-in is running doesn't support the needed minimum requirement set, you can get an attachment and its contents directly from Exchange instead. Select the tab for the applicable Exchange environment.

# [Exchange Online](#tab/exchange-online)

In Exchange Online environments, your add-in must perform the following steps to get attachments directly from Exchange.

1. Get an access token to [Microsoft Graph](/graph/overview).
1. Get the item ID of the applicable message or appointment.
1. Use Microsoft Graph to get the attachment and its properties.

Each step is covered in the following sections.

## Get an access token

Microsoft Graph provides access to users' Outlook mail data. Before your add-in can obtain data from Microsoft Graph, it must first get an access token for authorization. To get an access token, use nested app authentication (NAA). To learn more about NAA, see [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

## Get the item ID of the mail item

To get information about an attachment using Microsoft Graph, you need the item ID of the message or appointment that includes the attachment. Use the applicable Office JavaScript API to get the item ID.

- **Read mode**: Call [Office.context.mailbox.item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties). On non-mobile Outlook clients, because this property returns an ID formatted for Exchange Web Services (EWS), you must use the [Office.context.mailbox.convertToRestId](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-converttorestid-member(1)) method to convert the ID into a REST format that Microsoft Graph can use.

    ```javascript
    // Get the item ID of the current mail item in read mode and convert it into a REST format.
    const itemId = Office.context.mailbox.item.itemId;
    const restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    ```

- **Compose mode**: The method to get the item ID varies depending on whether the mail item has been saved as a draft.
  - If the item has been saved, call [Office.context.mailbox.item.getItemIdAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods).

    ```javascript
    // Get the item ID of the current mail item being composed.
    Office.context.mailbox.item.getItemIdAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error(result.error.message);
            return;
        }

        const itemId = result.value;
    });
    ```

    > [!TIP]
    > The `getItemIdAsync` method was introduced in Mailbox requirement set 1.8. If the Outlook client in which your add-in is running doesn't support Mailbox 1.8, use `Office.context.mailbox.item.saveAsync` instead as this method was introduced in Mailbox 1.3.

  - If the item hasn't been saved yet, call [Office.context.mailbox.item.saveAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) to initiate the save and get the item ID.

    ```javascript
    // Save the current mail item being composed to get its ID.
    Office.context.mailbox.item.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error(result.error.message);
            return;
        }

        const itemId = result.value;
    });
    ```

    > [!NOTE]
    > If your Outlook client is in cached mode, it may take some time for the saved item to sync to the server. Until the item is synced, using the item ID will return an error.

## Use Microsoft Graph

Once you've obtained an access token and the item ID of the mail item containing the attachment, you can now make a Microsoft Graph request. For information and examples on how to get an attachment using Microsoft Graph, see [Get attachment](/graph/api/attachment-get).

# [Exchange on-premises](#tab/exchange-on-prem)

In Exchange on-premises environments, your add-in must perform the following steps to get attachments directly from Exchange.

1. Get the callback token from the Exchange server.

1. Send the callback token and attachment information to the remote service.

1. Get the attachments from the Exchange server using the `ExchangeService.GetAttachments` method or the `GetAttachment` operation.

Each step is covered in detail in the following sections using code from the [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) sample.

> [!NOTE]
> The code in these examples has been shortened to emphasize the attachment information. The sample contains additional code for authenticating the add-in with the remote server and managing the state of the request.

## Get a callback token

The [Office.context.mailbox](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox) object provides the `getCallbackTokenAsync` method to get a token that the remote server can use to authenticate with the Exchange server. The following code shows a function in an add-in that starts the asynchronous request to get the callback token, and the callback function that gets the response. The callback token is stored in the service request object that is defined in the next section.

```js
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}

function attachmentTokenCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Couldn't get callback token: " + asyncResult.error.message);
    }
}
```

## Send attachment information to the remote service

The remote service that your add-in calls defines the specifics of how you should send the attachment information to the service. In this example, the remote service is a Web API application created by using Visual Studio. The remote service expects the attachment information in a JSON object.

The [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) property is used to provide the URL of Exchange Web Services (EWS) on the Exchange server that hosts the mailbox. This URL is then used by the remote service to call the [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation in a later step.

 The following code initializes an object that contains the attachment information.

```js
// Initialize a context object for the add-in.
// Set the fields that are used on the request
// object to default values.
 const serviceRequest = {
    attachmentToken: '',
    ewsUrl: Office.context.mailbox.ewsUrl,
    attachments: []
 };
```

The `Office.context.mailbox.item.attachments` property contains a collection of `AttachmentDetails` objects, one for each attachment to the item. In most cases, the add-in can pass just the attachment ID property of an `AttachmentDetails` object to the remote service. If the remote service needs more details about the attachment, you can pass all or part of the `AttachmentDetails` object. The following code defines a method that puts the entire `AttachmentDetails` array in the `serviceRequest` object and sends a request to the remote service.

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (let i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      const names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (let i = 0; i < response.attachmentNames.length; i++) {
        names += response.attachmentNames[i] + "<br />";
      }
      document.getElementById("names").innerHTML = names;
    } else {
      app.showNotification("Runtime error", response.message);
    }
  }).fail(function (status) {

  }).always(function () {
    $('.disable-while-sending').prop('disabled', false);
  })
}
```

## Get the attachments from the Exchange server

Your remote service can use either the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS Managed API method or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation to retrieve attachments from the server. The service application needs two objects to deserialize the JSON string into .NET Framework objects that can be used on the server. The following code shows the definitions of the deserialization objects.

```cs
namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```

### Use the EWS Managed API to get the attachments

If you use the [EWS Managed API](/exchange/client-developer/web-service-reference/ews-managed-api-reference-for-exchange) in your remote service, you can use the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, which will construct, send, and receive an EWS SOAP request to get the attachments. We recommend that you use the EWS Managed API because it requires fewer lines of code and provides a more intuitive interface for making calls to EWS. The following code makes one request to retrieve all the attachments, and returns the count and names of the attachments processed.

```cs
private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  // Create an ExchangeService object, set the credentials and the EWS URL.
  ExchangeService service = new ExchangeService();
  service.Credentials = new OAuthCredentials(request.attachmentToken);
  service.Url = new Uri(request.ewsUrl);

  var attachmentIds = new List<string>();

  foreach (AttachmentDetails attachment in request.attachments)
  {
    attachmentIds.Add(attachment.id);
  }

  // Call the GetAttachments method to retrieve the attachments on the message.
  // This method results in a GetAttachments EWS SOAP request and response
  // from the Exchange server.
  var getAttachmentsResponse =
    service.GetAttachments(attachmentIds.ToArray(),
                            null,
                            new PropertySet(BasePropertySet.FirstClassProperties,
                                            ItemSchema.MimeContent));

  if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
  {
    foreach (var attachmentResponse in getAttachmentsResponse)
    {
      attachmentNames.Add(attachmentResponse.Attachment.Name);

      // Write the content of each attachment to a stream.
      if (attachmentResponse.Attachment is FileAttachment)
      {
        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
        Stream s = new MemoryStream(fileAttachment.Content);
        // Process the contents of the attachment here.
      }

      if (attachmentResponse.Attachment is ItemAttachment)
      {
        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
        // Process the contents of the attachment here.
      }

      attachmentsProcessedCount++;
    }
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

### Use EWS to get the attachments

If you use EWS in your remote service, you need to construct a [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP request to get the attachments from the Exchange server. The following code returns a string that provides the SOAP request. The remote service uses the `String.Format` method to insert the attachment ID for an attachment into the string.

```cs
private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""https://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2016"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
```

Finally, the following method does the work of using an EWS `GetAttachment` request to get the attachments from the Exchange server. This implementation makes an individual request for each attachment, and returns the count of attachments processed. Each response is processed in a separate `ProcessXmlResponse` method, defined next.

```cs
private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  foreach (var attachment in request.attachments)
  {
    // Prepare a web request object.
    HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
    webRequest.Headers.Add("Authorization",
      string.Format("Bearer {0}", request.attachmentToken));
    webRequest.PreAuthenticate = true;
    webRequest.AllowAutoRedirect = false;
    webRequest.Method = "POST";
    webRequest.ContentType = "text/xml; charset=utf-8";

    // Construct the SOAP message for the GetAttachment operation.
    byte[] bodyBytes = Encoding.UTF8.GetBytes(
      string.Format(GetAttachmentSoapRequest, attachment.id));
    webRequest.ContentLength = bodyBytes.Length;

    Stream requestStream = webRequest.GetRequestStream();
    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
    requestStream.Close();

    // Make the request to the Exchange server and get the response.
    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

    // If the response is okay, create an XML document from the response
    // and process the request.
    if (webResponse.StatusCode == HttpStatusCode.OK)
    {
      var responseStream = webResponse.GetResponseStream();

      var responseEnvelope = XElement.Load(responseStream);

      // After creating a memory stream containing the contents of the
      // attachment, this method writes the XML document to the trace output.
      // Your service would perform it's processing here.
      if (responseEnvelope != null)
      {
        var processResult = ProcessXmlResponse(responseEnvelope);
        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

      }

      // Close the response stream.
      responseStream.Close();
      webResponse.Close();

    }
    // If the response is not OK, return an error message for the
    // attachment.
    else
    {
      var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
        "Error message: {1}.", attachment.name, webResponse.StatusDescription);
      attachmentNames.Add(errorString);
    }
    attachmentsProcessedCount++;
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

Each response from the `GetAttachment` operation is sent to the `ProcessXmlResponse` method. This method checks the response for errors. If it doesn't find any errors, it processes file attachments and item attachments. The `ProcessXmlResponse` method performs the bulk of the work to process the attachment.

```cs
// This method processes the response from the Exchange server.
// In your application the bulk of the processing occurs here.
private string ProcessXmlResponse(XElement responseEnvelope)
{
  // First, check the response for web service errors.
  var errorCodes = from errorCode in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                    select errorCode;
  // Return the first error code found.
  foreach (var errorCode in errorCodes)
  {
    if (errorCode.Value != "NoError")
    {
      return string.Format("Could not process result. Error: {0}", errorCode.Value);
    }
  }

  // No errors found, proceed with processing the content.
  // First, get and process file attachments.
  var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                        select fileAttachment;
  foreach(var fileAttachment in fileAttachments)
  {
    var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
    var fileData = System.Convert.FromBase64String(fileContent.Value);
    var s = new MemoryStream(fileData);
    // Process the file attachment here.
  }

  // Second, get and process item attachments.
  var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                        select itemAttachment;
  foreach(var itemAttachment in itemAttachments)
  {
    var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
    if (message != null)
    {
      // Process a message here.
      break;
    }
    var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
    if (calendarItem != null)
    {
      // Process calendar item here.
      break;
    }
    var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
    if (contact != null)
    {
      // Process contact here.
      break;
    }
    var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
    if (task != null)
    {
      // Process task here.
      break;
    }
    var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
    if (meetingMessage != null)
    {
      // Process meeting message here.
      break;
    }
    var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
    if (meetingRequest != null)
    {
      // Process meeting request here.
      break;
    }
    var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
    if (meetingResponse != null)
    {
      // Process meeting response here.
      break;
    }
    var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
    if (meetingCancellation != null)
    {
      // Process meeting cancellation here.
      break;
    }
  }

  return string.Empty;
}
```

---

## See also

- [Manage an item's attachments in a compose form in Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [Office Add-in sample: Single Sign-On(SSO) in an Outlook add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Use the Microsoft Graph API](/graph/use-the-api)
- [Explore the EWS Managed API, EWS, and web services in Exchange](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [Get started with EWS Managed API client applications](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
