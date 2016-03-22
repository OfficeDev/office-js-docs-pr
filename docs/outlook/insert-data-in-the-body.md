
# Insert data in the body when composing an appointment or message in Outlook

You can use the asynchronous methods ([Body.getAsync](../../reference/outlook/Body.md), [Body.getTypeAsync](../../reference/outlook/Body.md), [Body.prependAsync](../../reference/outlook/Body.md), [Body.setAsync](../../reference/outlook/Body.md) and [Body.setSelectedDataAsync](../../reference/outlook/Body.md)) to get the body type and insert data in the body of an appointment or message item that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in the section [Setting up Outlook add-ins for compose forms](../outlook/compose-scenario.md#mod_off15_CreatingForCompose_SettingUp) of [Create Outlook add-ins for compose forms](../outlook/compose-scenario.md).

In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting, you should always first verify the supported item format by calling  **getTypeAsync**, as you may need to take additional steps. The value that  **getTypeAsync** returns depends on the original item format, as well as the support of the device operating system and host to editing in HTML format (1). Then set the  _coercionType_ parameter of **prependAsync** or **setSelectedDataAsync** accordingly (2) to insert the data, as shown in the following table. If you don't specify an argument, **prependAsync** and **setSelectedDataAsync** assume the data to insert is in text format.



|**Data to insert**|**Item format returned by getTypeAsync**|**Use this coercionType**|
|:-----|:-----|:-----|
|Text|Text (1)|Text|
|HTML|Text (1)|Text (2)|
|Text|HTML|Text/HTML|
|HTML|HTML |HTML|

1.  On tablets and smartphones,  **getTypeAsync** returns **Office.MailboxEnums.BodyType.Text** if the operating system or host does not support editing an item, which was originally created in HTML, in HTML format.

2.  If your data to insert is HTML and  **getTypeAsync** returns a text type for that item, reorganize your data as text and insert it with **Office.MailboxEnums.BodyType.Text** as _coercionType_. If you simply insert the HTML data with a text coercion type, the host would display the HTML tags as text. If you attempt to insert the HTML data with  **Office.MailboxEnums.BodyType.Html** as _coercionType_, you will get an error.

In addition to  _coercionType_, as with most asynchronous methods in the JavaScript API for Office,  **getTypeAsync**,  **prependAsync** and **setSelectedDataAsync** take other optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../../docs/develop/asynchronous-programming-in-office-add-ins.md#AsyncProgramming_OptionalParameters) in [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## To insert data at the current cursor position


This section shows a code sample that uses  **getTypeAsync** to verify the body type of the item that is being composed, and then uses **setSelectedDataAsync** to insert data in the current cursor location.

You can pass a callback method and optional input parameters to  **getTypeAsync**, and get any status and results in the  _asyncResult_ output parameter. If the method succeeds, you can get the type of the item body in the [AsyncResult.value](../../reference/shared/asyncresult.status.md) property, which is either "text" or "html".

You must pass a data string as an input parameter to  **setSelectedDataAsync**. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned above, you can optionally specify the type of the data to be inserted in the  _coercionType_ parameter. In addition, you can provide a callback method and any of its parameters as optional input parameters.

If the user hasn't placed the cursor in the item body,  **setSelectedDataAsync** inserts the data at the top of the body. If the user has selected text in the item body, **setSelectedDataAsync** replaces the selected text by the data you specify. Note that **setSelectedDataAsync** can fail if the user is simultaneously changing the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.

This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## To insert data at the beginning of the item body


Alternatively, you can use  **prependAsync** to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, **prependAsync** and **setSelectedDataAsync** behave in similar ways:


- If you are prepending HTML data in a message body, you should first check for the type of the message body to avoid prepending HTML data to a message in text format.
    
- Provide the following as input parameters to  **prependAsync**: a data string in either text or HTML format, and optionally the format of the data to be inserted, a callback method and any of its parameters.
    
- The maximum number of characters you can prepend at one time is 1,000,000 characters.
    
The following JavaScript code is part of a sample add-in that is activated in compose forms of appointments and messages. The sample calls  **getTypeAsync** to verify the type of the item body, inserts HTML data to the top of the item body if the item is an appointment or HTML message, otherwise inserts the data in text format.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Additional resources



- [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Get and set Outlook item data in read or compose forms](../outlook/item-data.md)
    
- [Create Outlook add-ins for compose forms](../outlook/compose-scenario.md)
    
- [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Get, set, or add recipients when composing an appointment or message in Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Get or set the subject when composing an appointment or message in Outlook](../outlook/get-or-set-the-subject.md)
    
- [Get or set the location when composing an appointment in Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Get or set the time when composing an appointment in Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
