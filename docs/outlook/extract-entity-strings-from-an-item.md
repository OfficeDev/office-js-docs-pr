
# Extract entity strings from an Outlook item

This article describes how to create a  **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation. The supported entities include:

- Address: A United States postal address, that has at least a subset of the elements of a street number, street name, city, state, and zip code.
    
- Contact: A person's contact information, in the context of other entities such as an address or business name.
    
- Email address: An SMTP email address.
    
- Meeting suggestion: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.
    
- Phone number: A North American phone number.
    
- Task suggestion: A task suggestion, typically expressed in an actionable phrase.
    
- URL.
    
Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.
Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item. 

The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.

## XML manifest


The entities add-in has two activation rules joined by a logical OR operation. 


```XML
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

These rules specify that Outlook should activate this add-in when the currently selected item in the Reading Pane or read inspector is an appointment or message (including an email message, or meeting request, response, or cancellation).

The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.




```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## HTML implementation


The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#olowa15conagave_entitiesJSImplementation). The JavaScript file includes the event handlers for each of the buttons.

Note that all Outlook add-ins must include office.js. The HTML file that follows includes version 1.1 of office.js on the CDN. 




```HTML
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" 
        href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/Office.js"
        type="text/javascript"></script>
    <script type="text/javascript" 
        src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## Style sheet


The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.


```
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## JavaScript implementation


The remaining sections describe how this sample (default_entities.js file) extracts well-known entities from the subject and body of the message or appointment that the user is viewing. 


## Extracting entities upon initialization


Upon the [Office.initialize](../../reference/shared/office.initialize.md) event, the entities add-in calls the [getEntities](../../reference/outlook/Office.context.mailbox.item.md) method of the current item. The **getEntities** method returns the global variable `_MyEntities` an array of instances of supported entities. The following is the related JavaScript code.


```
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## Extracting addresses


When the user clicks the  **Get Addresses** button, the `myGetAddresses` event handler obtains an array of addresses from the [addresses](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any address was extracted. Each extracted address is stored as a string in the array. `myGetAddresses` forms a local HTML string in .mdText` to display the list of extracted addresses. The following is the related JavaScript code.


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## Extracting contact information


When the user clicks the  **Get Contact Information** button, the `myGetContacts` event handler obtains an array of contacts together with their information from the [contacts](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any was extracted. Each extracted contact is stored as a [Contact](../../reference/outlook/simple-types.md) object in the array. `myGetContacts` obtains further data about each contact. Note that the context determines whether Outlook can extract a contact from an item - a signature at the end of an email message, or at least some of the following information would have to exist in the vicinity of the contact:


- The string representing the contact's name from the [Contact.personName](../../reference/outlook/simple-types.md) property.
    
- The string representing the company name associated with the contact from the [Contact.businessName](../../reference/outlook/simple-types.md) property.
    
- The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](../../reference/outlook/simple-types.md) property. Each telephone number is represented by a [PhoneNumber](../../reference/outlook/simple-types.md) object.
    
- For each  **PhoneNumber** member in the telephone numbers array, the string representing the telephone number from the [PhoneNumber.phoneString](../../reference/outlook/simple-types.md) property.
    
- The array of URLs associated with the contact from the [Contact.urls](../../reference/outlook/simple-types.md) property. Each URL is represented as a string in an array member.
    
- The array of email addresses associated with the contact from the [Contact.emailAddresses](../../reference/outlook/simple-types.md) property. Each email address is represented as a string in an array member.
    
- The array of postal addresses associated with the contact from the [Contact.addresses](../../reference/outlook/simple-types.md) property. Each postal address is represented as a string in an array member.
    
 `myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## Extracting email addresses


When the user clicks the  **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Extracting meeting suggestions


When the user clicks the  **Get Meeting Suggestions** button, the `myGetMeetingSuggestions` event handler obtains an array of meeting suggestions from the [meetingSuggestions](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any was extracted.


 >**Note**  Only messages but not appointments support the  **MeetingSuggestion** entity type.

Each extracted meeting suggestion is stored as a [MeetingSuggestion](../../reference/outlook/simple-types.md) object in the array. `myGetMeetingSuggestions` obtains further data about each meeting suggestion:


- The string that was identified as a meeting suggestion from the [MeetingSuggestion.meetingString](../../reference/outlook/simple-types.md) property.
    
- The array of meeting attendees from the [MeetingSuggestion.attendees](../../reference/outlook/simple-types.md) property. Each attendee is represented by an [EmailUser](../../reference/outlook/simple-types.md) object.
    
- For each attendee, the name from the [EmailUser.displayName](../../reference/outlook/simple-types.md) property.
    
- For each attendee, the SMTP address from the [EmailUser.emailAddress](../../reference/outlook/simple-types.md) property.
    
- The string representing the location of the meeting suggestion from the [MeetingSuggestion.location](../../reference/outlook/simple-types.md) property.
    
- The string representing the subject of the meeting suggestion from the [MeetingSuggestion.subject](../../reference/outlook/simple-types.md) property.
    
- The string representing the start time of the meeting suggestion from the [MeetingSuggestion.start](../../reference/outlook/simple-types.md) property.
    
- The string representing the end time of the meeting suggestion from the [MeetingSuggestion.end](../../reference/outlook/simple-types.md) property.
    
 `myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.




```
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
           meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
            meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Extracting phone numbers


When the user clicks the  **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](../../reference/outlook/simple-types.md) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:


- The string representing the kind of phone number, for example, home phone number, from the [PhoneNumber.type](../../reference/outlook/simple-types.md) property.
    
- The string representing the actual phone number from the [PhoneNumber.phoneString](../../reference/outlook/simple-types.md) property.
    
- The string that was originally identified as the phone number from the [PhoneNumber.originalPhoneString](../../reference/outlook/simple-types.md) property.
    
 `myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.




```
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Extracting task suggestions


When the user clicks the  **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](../../reference/outlook/simple-types.md) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:


- The string that was originally identified a task suggestion from the [TaskSuggestion.taskString](../../reference/outlook/simple-types.md) property.
    
- The array of task assignees from the [TaskSuggestion.assignees](../../reference/outlook/simple-types.md) property. Each assignee is represented by an [EmailUser](../../reference/outlook/simple-types.md) object.
    
- For each assignee, the name from the [EmailUser.displayName](../../reference/outlook/simple-types.md) property.
    
- For each assignee, the SMTP address from the [EmailUser.emailAddress](../../reference/outlook/simple-types.md) property.
    
 `myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.




```
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Extracting URLs


When the user clicks the  **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](../../reference/outlook/simple-types.md) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Clearing displayed entity strings


Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## JavaScript listing


The following is the complete listing of the JavaScript implementation.


```
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Additional resources



- [Create Outlook add-ins for read forms](../outlook/read-scenario.md)
    
- [Match strings in an Outlook item as well-known entities](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [item.getEntities method](../../reference/outlook/Office.context.mailbox.item.md)
    
