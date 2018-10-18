---
title: Get or set the subject in an Outlook add-in | Microsoft Docs
description: Learn how to get or set the subject of a message or appointment in an Outlook add-in.
author: jasonjoh
ms.topic: article
ms.technology: office-add-ins
ms.date: 06/13/2017
ms.author: jasonjoh
---

# Get or set the subject when composing an appointment or message in Outlook

The JavaScript API for Office provides asynchronous methods ([subject.getAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Subject#getasync-options--callback-) and [subject.setAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Subject#setasync-subject--options--callback-)) to get and set the subject of an appointment or message that the user is composing. These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms.

The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:

```js
item.subject
```

But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the asynchronous method **getAsync** to get the subject, as shown below:

```js
item.subject.getAsync
```

The **subject** property is available for write access in only compose forms and not in read forms.

As with most asynchronous methods in the JavaScript API for Office, **getAsync** and **setAsync** take optional input parameters. For more information about specifying these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/asynchronous-programming-in-office-add-ins).


## Get the subject

This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

To use **item.subject.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](https://docs.microsoft.com/javascript/api/office/office.asyncresult#value) property.


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Set the subject


This section shows a code sample that sets the subject of the appointment or message that the user is composing. Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message.

To use **item.subject.setAsync**, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the  _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified subject string as plain text, overwriting any existing subject for that item.

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## See also

- [Get and set item data in a compose form in Outlook](get-and-set-item-data-in-a-compose-form.md)   
- [Get and set Outlook item data in read or compose forms](item-data.md)    
- [Create Outlook add-ins for compose forms](compose-scenario.md)    
- [Asynchronous programming in Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/asynchronous-programming-in-office-add-ins)    
- [Get, set, or add recipients when composing an appointment or message in Outlook](get-set-or-add-recipients.md)  
- [Insert data in the body when composing an appointment or message in Outlook](insert-data-in-the-body.md)   
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Get or set the time when composing an appointment in Outlook](get-or-set-the-time-of-an-appointment.md)
    
