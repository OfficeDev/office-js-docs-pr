
# Add and remove attachments to an item in a compose form in Outlook

You can use the [addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) and [addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) methods to attach a file and an Outlook item respectively to the item that the user is composing. Both are asynchronous methods, which means execution can go on without waiting for the add-attachment action to complete. Depending on the original location and size of the attachment being added, the add-attachment asynchronous call may take a while to complete. If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method. This callback method is optional and is invoked when the uploading of the attachment is complete. The callback method takes an [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.md) object as an output parameter that provides any status, error, and returned value from the add-attachment action. If the callback requires any extra parameters, you can specify them in the optional _options.aysncContext_ parameter. _options.asyncContext_ can be of any type that your callback method expects.

For example, you can define  _options.asyncContext_ as a JSON object that contains one or more key-value pairs, with the ':' character separating a key and the value, and a ',' separating one key-value pair from another. You can find more examples about [passing optional parameters to asynchronous methods](../../docs/develop/asynchronous-programming-in-office-add-ins.md#AsyncProgramming_OptionalParameters) in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md). The following example shows how to use the  **asyncContext** parameter to pass 2 arguments to a callback method:




```js
{ asyncContext: { var1: 1, var2: 2} }
```

You can check for success or error of an asynchronous method call in the callback method using the  **status** and **error** properties of the **AsyncResult** object. If the attaching completes successfully, you can use the **AsyncResult.value** property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.


 >**Note**  As a best practice, you should use the attachment ID to remove an attachment only if the same add-in has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment ID is valid only within the same session. A session is over when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.


## Attaching a file

You can attach a file to a message or appointment in a compose form by using the  **addFileAttachmentAsync** method and specifying the URI of the file. If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter. Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.

The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes  **asyncResult** as a parameter, checks for the attaching status, and gets the attachment ID if the attaching succeeds.




```js
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Attaching an Outlook item

You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the  **addItemAttachmentAsync** method. You can get the EWS ID of an email, calendar, contact or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) method and accessing the EWS operation [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx). The [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md) property also provides the EWS ID of an existing item in a read form.

The following JavaScript function,  `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed. The function takes as an argument the EWS ID of the item that is to be attached. If the attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.




```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


 >**Note**  You can use a compose add-in to attach an instance of a recurring appointment in Outlook Web App or OWA for Devices. However, in a supporting Outlook rich client, attempting to attach an instance would result in attaching the recurring series (the master appointment).


## Removing an attachment


You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID and using the [removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) method. You should remove only attachments that the same add-in has added in the same session. You should make sure the attachment ID corresponds to a valid attachment, or the method will return an error. Similar to the **addFileAttachmentAsync** and **addItemAttachmentAsync** methods, **removeAttachmentAsync** is an asynchronous method. You should provide a callback method to check for the status and any error by using the **AsyncResult** output parameter object. You can also pass any additional parameters to the callback method by using the optional **asyncContext** parameter, which is a JSON object of key-value pairs.

The following JavaScript function,  `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed. The function takes as an argument the ID of the attachment to be removed. You can obtain the ID of an attachment after a successful  **addFileAttachmentAsync** or **addItemAttachmentAsync** method call, and store it for a subsequent **removeAttachmentAsync** method call.




```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## Tips for adding and removing attachments


If your compose add-in adds and removes attachments, structure your code such that you pass a valid attachment ID to the remove-attachment call, and handle the case when  **AsyncResult.error** returns **InvalidAttachmentId**. Depending on the location and size of an attachment, attaching a file or item can take some time to complete. The following example contains a call to  **addFileAttachmentAsync**,  `write`, and  **removeAttachmentAsync**. You might think that the calls would execute sequentially one after another.


```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

Even though  **addFileAttachmentAsync** starts before **removeAttachmentAsync**, because  **addFileAttachmentAsync** is asynchronous, the `write` and **removeAttachmentAsync** calls can start before **addFileAttachmentAsync** completes. When this happens, `attachmentID` remains **undefined**, and you will get an error for the  **removeAttachmentAsync** call, as in the following output:




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

One way to avoid this is to check that  `attachmentID` is defined before calling **removeAttachmentAsync**. Another way is to initiate the  **removeAttachmentAsync** call from within the callback method of **addFileAttachmentAsync**, as shown in the following example:




```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

The following is an example of the output:




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

Notice that the callback for  **removeAttachmentAsync** is nested inside the callback for **addFileAttachmentAsync**. Because  **addFileAttachmentAsync** and **removeAttachmentAsync** are asynchronous, the last line in the callback for **addFileAttachmentAsync** can get executed before the callback for **removeAttachmentAsync** completes.


## Additional resources



- [Create Outlook add-ins for compose forms](../outlook/compose-scenario.md)
    
- [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


