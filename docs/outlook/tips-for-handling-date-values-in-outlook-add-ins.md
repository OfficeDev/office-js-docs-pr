
# Tips for handling date values in Outlook add-ins

The JavaScript API for Office uses the JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) object for most of the storage and retrieval of dates and times. That **Date** object provides methods such as [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp), and [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp), which return the requested date or time value according to Universal Coordinated Time (UTC) time.<br/><br/>
The **Date** object also provides other methods such as [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp), and [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), which return the requested date or time according to "local time".<br/><br/>
The concept of "local time" is largely determined by the browser and operating system on the client computer. For instance, on most browsers running on a Windows-based client computer, a JavaScript call to **getDate**, returns a date based on the time zone set in Windows on the client computer.<br/><br/>
The following example creates a  **Date** object <code>myLocalDate</code> in local time, and calls **toUTCString** to convert that date to a date string in UTC.




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

While you can use the JavaScript  **Date** object to get a date or time value based on UTC or the client computer time zone, the **Date** object is limited in one respect - it does not provide methods to return a date or time value for any other specific time zone. For example, if your client computer is set to be on Eastern Standard Time (EST), there is no **Date** method that allows you to get the hour value other than in EST or UTC, such as Pacific Standard Time (PST).


## Date-related features for Outlook add-ins


The afore-mentioned JavaScript limitation has an implication for you, when you use the JavaScript API for Office to handle date or time values in Outlook add-ins that run in an Outlook rich client, and in Outlook Web App or OWA for Devices.


### Time zones for Outlook clients

For clarity, let's define the time zones in question.



|**Time zone**|**Description**|
|:-----|:-----|
|Client computer time zone|This is set on the operating system of the client computer. Most browsers use the client computer time zone to display date or time values of the JavaScript  **Date** object.<br/><br/>An Outlook rich client uses this time zone to display date or time values in the user interface. <br/><br/>For example, on a client computer running Windows, Outlook uses the time zone set on Windows as the local time zone. On the Mac, if the user changes the time zone on the client computer, Outlook for Mac would prompt the user to update the time zone in Outlook as well.|
|Exchange Admin Center (EAC) time zone|The user sets this time zone value (and the preferred language) when he or she logs on to Outlook Web App or OWA for Devices the first time. <br/><br/>Outlook Web App and OWA for Devices use this time zone to display date or time values in the user interface.|
Because an Outlook rich client uses the client computer time zone, and the user interface of Outlook Web App and OWA for Devices uses the EAC time zone, the local time for the same add-in installed for the same mailbox can be different when running in an Outlook rich client and in Outlook Web App or OWA for Devices. As an Outlook add-in developer, you should appropriately input and output date values so that those values are always consistent with the time zone that the user expects on the corresponding client.


### Date-related API

The following are the properties and methods in the JavaScript API for Office that support date-related features.reference/outlook/Office.context.mailbox.item.md



**API member**|**Time zone representation**|**Example in an Outlook rich client**|**Example in Outlook Web App or OWA for Devices**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|In an Outlook rich client, this property returns the client computer time zone. In Outlook Web App and OWA for Devices, this property returns the EAC time zone. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) and [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Each of these properties returns a JavaScript  **Date** object. This **Date** value is UTC-correct, as shown in the following example - `myUTCDate` has the same value in an Outlook rich client, Outlook Web App and OWA for Devices.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>However, calling  `myDate.getDate` returns a date value in the client computer time zone, which is consistent with the time zone used to display date times values in the Outlook rich client interface, but may be different from the EAC time zone that Outlook Web App and OWA for Devices use in its user interface.|If the item is created at 9am UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` returns 4am EST.<br/><br/>If the item is modified at 11am UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` returns 6am EST.|If the item creation time is 9am UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` returns 4am EST.<br/><br/>If the item is modified at 11am UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` returns 6am EST.<br/><br/>Notice that if you want to display the creation or modification time in the user interface, you would want to first convert the time to PST to be consistent with the rest of the user interface.
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|Each of the  _Start_ and _End_ parameters requires a JavaScript **Date** object. The arguments should be UTC-correct regardless of the time zone used in the user interface of an Outlook rich client, Outlook Web App or OWA for Devices.|If the start and end times for the appointment form are 9am UTC and 11am UTC, then you should assure that the `start` and `end` arguments are UTC-correct, which means:<br/><br/><ul><li>`start.getUTCHours` returns 9am UTC</li><li>`end.getUTCHours` returns 11am UTC</li></ul>|If the start and end times for the appointment form are 9am UTC and 11am UTC, then you should assure that the `start` and `end` arguments are UTC-correct, which means:<br/><br/><ul><li>`start.getUTCHours` returns 9am UTC</li><li>`end.getUTCHours` returns 11am UTC</li></ul>

## Helper methods for date-related scenarios


As described in the preceding sections, because the "local time" for a user in Outlook Web App or OWA for Devices can be different on an Outlook rich client, but the JavaScript  **Date** object supports converting to only the client computer time zone or UTC, the JavaScript API for Office provides two helper methods: [Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) and [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md). <br/><br/>
These helper methods take care of any need to handle date or time differently for the following two date-related scenarios, in an Outlook rich client, Outlook Web App and OWA for Devices, thus reinforcing "write-once" for different clients of your add-in.


### Scenario A: displaying item creation or modified time

If you are displaying the item creation time (**Item.dateTimeCreated**) or modification time (**Item.dateTimeModified**) in the user interface, first use  **convertToLocalClientTime** to convert the **Date** object provided by these properties to obtain a dictionary representation in the appropriate local time. Then display the parts of the dictionary date. The following is an example of this scenario:


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Note that  **convertToLocalClientTime** takes care of the difference between an Outlook rich client, and Outlook Web App or OWA for Devices:


- If  **convertToLocalClientTime** detects the current host is a rich client, the method converts the **Date** representation to a dictionary representation in the same client computer time zone, consistent with the rest of the rich client user interface.
    
- If  **convertToLocalClientTime** detects the current host is Outlook Web App or OWA for Devices, the method converts the UTC-correct **Date** representation to a dictionary format in the EAC time zone, consistent with the rest of the Outlook Web App or OWA for Devices user interface.
    

### Scenario B: displaying start and end dates in a new appointment form

If you are obtaining as input different parts of a date-time value represented in the local time, and would like to provide this dictionary input value as a start or end time in an appointment form, first use the  **convertToUtcClientTime** helper method to convert the dictionary value to a UTC-correct **Date** object.<br/><br/>
In the following example, assume  `myLocalDictionaryStartDate` and `myLocalDictionaryEndDate` are date-time values in dictionary format that you have obtained from the user. These values are based on the local time, dependent on the host application.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

The resultant values,  `myUTCCorrectStartDate` and `myUTCCorrectEndDate`, are UTC-correct. Then pass these  **Date** objects as arguments for the _Start_ and _End_ parameters of the **Mailbox.displayNewAppointmentForm** method to display the new appointment form.<br/><br/>
Note that  **convertToUtcClientTime** takes care of the difference between an Outlook rich client, and Outlook Web App or OWA for Devices:


- If  **convertToUtcClientTime** detects the current host is an Outlook rich client, the method simply converts the dictionary representation to a **Date** object. This **Date** object is UTC-correct, as expected by **displayNewAppointmentForm**.
    
- If  **convertToUtcClientTime** detects the current host is Outlook Web App or OWA for Devices, the method converts the dictionary format of the date and time values expressed in the EAC time zone to a **Date** object. This **Date** object is UTC-correct, as expected by **displayNewAppointmentForm**.
    

## Additional resources



- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md)
    


