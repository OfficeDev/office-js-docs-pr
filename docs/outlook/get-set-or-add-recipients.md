
# Get, set, or add recipients when composing an appointment or message in Outlook


The JavaScript API for Office provides asynchronous methods ([Recipients.getAsync](../../reference/outlook/Recipients.md), [Recipients.setAsync](../../reference/outlook/Recipients.md), or [Recipients.addAysnc](../../reference/outlook/Recipients.md)) to respectively get, set, or add recipients in a compose form of an appointment or message. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in the section [Setting up Outlook add-ins for compose forms](../outlook/compose-scenario.md#mod_off15_CreatingForCompose_SettingUp) of [Create Outlook add-ins for compose forms](../outlook/compose-scenario.md).

Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) and [requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) for appointments, and [cc](../../reference/outlook/Office.context.mailbox.item.md), and  [to](../../reference/outlook/Office.context.mailbox.item.md) for messages. In a read form, you can access the property directly from the parent object, such as:




```js
item.cc
```

But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method  **getAsync** to get these properties, as in the following example:




```js
item.cc.getAsync
```

These properties are available for write access in only compose forms and not read forms.

As with most asynchronous methods in the JavaScript API for Office,  **getAsync**,  **setAsync**, and  **addAsync** take optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../../docs/develop/asynchronous-programming-in-office-add-ins.md#AsyncProgramming_OptionalParameters) in [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## To get recipients


This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below. 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

In the JavaScript API for Office, because the properties that represent the recipients of an appointment ( **optionalAttendees** and **requiredAttendees**) are different from those of a message ([bcc](../../reference/outlook/Office.context.mailbox.item.md),  **cc**, and  **to**), you should first use the [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) property to identify whether the item being composed is an appointment or message. In compose mode, all these properties of appointments and messages are [Recipients](../../reference/outlook/Recipients.md) objects, so you can then apply the asynchronous method, **Recipients.getAsync**, to get the corresponding recipients. 

To use  **getAsync**, provide a callback method to check for the status, results, and any error returned by the asynchronous  **getAsync** call. You can provide any arguments to the callback method using the optional _asyncContext_ parameter. The callback method returns an _asyncResult_ output parameter. You can use the **status** and **error** properties of the [AsyncResult](../../reference/outlook/simple-types.md) parameter object to check for status and any error messages of the asynchronous call, and the **value** property to get the actual recipients. Recipients are represented as an array of [EmailAddressDetails](../../reference/outlook/simple-types.md) objects.

Note that because the  **getAsync** method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback method when the asynchronous call has successfully completed.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## To set recipients


This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user. Setting recipients overwrites any existing recipients. Similar to the previous example that gets recipients in a compose form, this example assumes that the add-in is activated in compose forms for appointments and messages. This example first verifies if the composed item is an appointment or message, so to apply the asynchronous method,  **Recipients.setAsync**, on the appropriate properties that represent recipients of the appointment or message.

When calling  **setAsync**, provide an array as input argument for the  _recipients_ parameter, in one of the following formats:


- An array of strings that are SMTP addresses.
    
- An array of dictionaries, each containing a display name and email address, as shown in the following code sample.
    
- An array of  **EmailAddressDetails** objects, similar to the one returned by the **getAsync** method.
    
You can optionally provide a callback method as an input argument to the  **setAsync** method, to make sure any code that depends on successfully setting the recipients would execute only when that happens. You can also provide any arguments for the callback method using the optional _asyncContext_ parameter. If you use a callback method, you can access an _asyncResult_ output parameter, and use the **status** and **error** properties of the **AsyncResult** parameter object to check for status and any error messages of the asynchronous call.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## To add recipients


If you do not want to overwrite any existing recipients in an appointment or message, instead of using  **Recipients.setAsync**, you can use the  **Recipients.addAsync** asynchronous method to append recipients. **addAsync** works similarly as **setAsync** in that it requires a _recipients_ input argument. You can optionally provide a callback method, and any arguments for the callback using the asyncContext parameter. You can then check the status, result, and any error of the asynchronous **addAsync** call by using the _asyncResult_ output parameter of the callback method. The following example checks if the item being composed is an appointment, and appends two required attendees to the appointment.


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## Additional resources



- [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Get and set Outlook item data in read or compose forms](../outlook/item-data.md)
    
- [Create Outlook add-ins for compose forms](../outlook/compose-scenario.md)
    
- [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Get or set the subject when composing an appointment or message in Outlook](../outlook/get-or-set-the-subject.md)
    
- [Insert data in the body when composing an appointment or message in Outlook](../outlook/insert-data-in-the-body.md)
    
- [Get or set the location when composing an appointment in Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Get or set the time when composing an appointment in Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
