
# Get attachments of an Outlook item from the server

An Outlook add-in cannot pass the attachments of a selected item directly to the remote service that runs on your server. Instead, the add-in can use the attachments API to send information about the attachments to the remote service. The service can then contact the Exchange server directly to retrieve the attachments.

To send attachment information to the remote service, you use the following properties and function:


- [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) property -- Provides the URL of Exchange Web Services (EWS) on the Exchange server that hosts the mailbox. Your service uses this URL to call the [ExchangeService.GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx)[EWS Managed API](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx) method or the [GetAttachment](http://msdn.microsoft.com/en-us/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) EWS operation.
    
- [Office.context.mailbox.item.attachments](../../reference/outlook/Office.context.mailbox.item.md) property -- Gets an array of [AttachmentDetails](../../reference/outlook/simple-types.md) objects, one for each attachment to the item.
    
- [Office.context.mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) function -- Makes an asynchronous call to the Exchange server that hosts the mailbox to get a callback token that the server sends back to the Exchange server to authenticate a request for an attachment.
    

## Using the attachments API


To use the attachments API to get attachments from an Exchange mailbox, perform the following steps: 


1. Show the add-in when the user is viewing a message or appointment that contains an attachment.
    
2. Get the callback token from the Exchange server.
    
3. Send the callback token and attachment information to the remote service.
    
4. Get the attachments from the Exchange server by using the  **ExchangeService.GetAttachments** method or the **GetAttachment** operation.
    
Each of these steps is covered in detail in the following sections using code from the [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) sample.


 >**Note**  The code in these examples has been shortened to emphasize the attachment information. The sample contains additional code for authenticating the add-in with the remote server and managing the state of the request.


### Activate the add-in


You can use an [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) rule in the add-in manifest file to show your add-in when the selected item has attachments, as shown in the following example.


```XML
<Rule xsi:type="ItemHasAttachment" />
```


### Get a callback token


The [Office.context.mailbox](../../reference/outlook/Office.context.mailbox.md) object provides the **getCallbackTokenAsync** function to get a token that the remote server can use to authenticate with the Exchange server. The following code shows a function in an add-in that starts the asynchronous request to get the callback token, and the callback function that gets the response. The callback token is stored in the service request object that is defined in the next section.


```
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "" {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
};
function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
};
```


### Send attachment information to the remote service


The remote service that your add-in calls defines the specifics of how you should send the attachment information to the service. In this example, the remote service is Web API application created by using Visual Studio 2013. The remote service expects the attachment information in a JSON object. The following code initializes an object that contains the attachment information.


```
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
serviceRequest = new Object();
serviceRequest.attachmentToken = "";
serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
serviceRequest.attachments = new Array();
```

The  `Office.context.mailbox.item.attachments` property contains a collection of **AttachmentDetails** objects, one for each attachment to the item. In most cases, the add-in can pass just the attachment ID property of an **AttachmentDetails** object to the remote service. If the remote service needs more details about the attachment, you can pass all or part of the **AttachmentDetails** object. The following code defines a method that puts the entire **AttachmentDetails** array in the `serviceRequest` object and sends a request to the remote service.




```js
    function makeServiceRequest() {
      // Format the attachment details for sending.
      for (var i = 0; i < mailbox.item.attachments.length; i++) {
        serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i].$0_0));
      }

      $.ajax({
        url: '../../api/Default',
        type: 'POST',
        data: JSON.stringify(serviceRequest),
        contentType: 'application/json;charset=utf-8'
      }).done(function (response) {
        if (!response.isError) {
          var names = "<h2>Attachments processed using " +
                        serviceRequest.service +
                        ": " +
                        response.attachmentsProcessed +
                        "</h2>";
          for (i = 0; i < response.attachmentNames.length; i++) {
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
    };

```


### Get the attachments from the Exchange server


Your remote service can use either the [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) EWS Managed API method or the [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) EWS operation to retrieve attachments from the server. The service application needs two objects to deserialize the JSON string into .NET Framework objects that can be used on the server. The following code shows the definitions of the deserialization objects.


```C#



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


#### Use the EWS Managed API to get the attachments

If you use the [EWS Managed API](http://go.microsoft.com/fwlink/?LinkID=255472) in your remote service, you can use the [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) method, which will construct, send, and receive an EWS SOAP request to get the attachments. We recommend that you use the EWS Managed API because it requires fewer lines of code and provides a more intuitive interface for making calls to EWS. The following code makes one request to retrieve all the attachments, and returns the count and names of the attachments processed.


```C#
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


#### Use EWS to get the attachments

If you use EWS in your remote service, you need to construct a [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) SOAP request to get the attachments from the Exchange server. The following code returns a string that provides the SOAP request. The remote service uses the **String.Format** method to insert the attachment ID for an attachment into the string.


```C#
    private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
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

Finally, the following method does the work of using an EWS  **GetAttachment** request to get the attachments from the Exchange server. This implementation makes an individual request for each attachment, and returns the count of attachments processed. Each response is processed in a separate **ProcessXmlResponse** method, defined next.




```C#
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

        // If the response is okay, create an XML document from the reponse
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

Each response from the  **GetAttachment** operation is sent to the **ProcessXmlResponse** method. This method checks the response for errors. If it doesn't find any errors, it processes file attachments and item attachments. The **ProcessXmlResponse** method performs the bulk of the work to process the attachment.




```C#
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


## Additional resources



- [Create Outlook add-ins for read forms](../outlook/read-scenario.md)
    
- [Explore the EWS Managed API, EWS, and web services in Exchange](http://msdn.microsoft.com/library/0bc6f81d-cc10-42b0-ba5d-6f22ff55d51c%28Office.15%29.aspx)
    
- [Get started with EWS Managed API client applications](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx)
    
- [Outlook-Power-Hour_Code-Samples](https://github.com/OfficeDev/Outlook-Power-Hour-Code-Samples):  `MyAttachments` and `AttachmentsDemo`
    
