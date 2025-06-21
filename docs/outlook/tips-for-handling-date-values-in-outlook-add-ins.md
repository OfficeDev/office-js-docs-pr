---
title: Handle date values in Outlook add-ins
description: The Office JavaScript API uses the JavaScript Date object for most of the storage and retrieval of dates and times. 
ms.date: 04/17/2025
ms.localizationpriority: medium
---

# Handle date values in Outlook add-ins

The Office JavaScript API uses the JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) object for most of the storage and retrieval of dates and times.

The `Date` object provides methods such as [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp), and [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), which return the requested date or time value according to Universal Coordinated Time (UTC) time.

It also provides other methods such as [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp), and [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), which return the requested date or time according to "local time".

The concept of "local time" is largely determined by the browser and operating system on the client computer. For instance, on most browsers running on a Windows-based client computer, a JavaScript call to `getDate`, returns a date based on the time zone set in Windows on the client computer.

The following example creates `myLocalDate`, which is a `Date` object in local time, and calls `toUTCString` to convert that date to a date string in UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
const myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

While you can use the JavaScript `Date` object to get a date or time value based on UTC or the client computer time zone, the `Date` object is limited in one respect - it doesn't provide methods to return a date or time value for any other specific time zone. For example, if your client computer is set to be on Eastern Standard Time (EST), there is no `Date` method that allows you to get the hour value other than in EST or UTC, such as Pacific Standard Time (PST).

## Date-related features for Outlook add-ins

The aforementioned JavaScript limitation has an implication for you when you use the Office JavaScript API to handle date or time values in add-ins that run in Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) or classic), on Mac, on the web, or on mobile devices.

### Time zones for Outlook clients

For clarity, let's define the time zones in question.

|Time zone|Description|
|:-----|:-----|
|Client computer time zone|This is set on the operating system of the client computer. Most browsers use the client computer time zone to display date or time values of the JavaScript `Date` object.<br/><br/>Outlook on Windows (classic) and Outlook on Mac use this time zone to display date or time values in the user interface. <br/><br/>For example, on a client computer running Windows, classic Outlook uses the time zone set on Windows as the local time zone. On the Mac, if the user changes the time zone on the client computer, Outlook on Mac would prompt the user to update the time zone in Outlook as well.|
|Exchange Admin Center (EAC) time zone|The user sets this time zone value (and the preferred language) when they log on to Outlook on the web, new Outlook on Windows, or mobile devices the first time. <br/><br/>Outlook on the web, new Outlook on Windows, and mobile devices use this time zone to display date or time values in the user interface.|

Because Outlook on Windows (classic) and Outlook on Mac use the client computer time zone, and the user interface of Outlook on the web, new Outlook on Windows, and mobile devices uses the EAC time zone, the local time for the same add-in installed on the same mailbox can be different depending on which platform your Outlook client is running. As an Outlook add-in developer, you should appropriately input and output date values so that those values are always consistent with the time zone that the user expects on the corresponding client.

### Date-related API

The following are the properties and methods in the Office JavaScript API that support date-related features.

|API member|Time zone representation|Example in Outlook on the web, on new Windows client, and on mobile devices|Example in Outlook on Windows (classic) and on Mac|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|In Outlook on Windows (classic) and Outlook on Mac, this property returns the client computer time zone. In Outlook on the web, on mobile devices, and in new Outlook on Windows, this property returns the EAC time zone.|PST|EST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Each of these properties returns a JavaScript `Date` object. This `Date` value is UTC-correct, as shown in the following example - `myUTCDate` has the same value in Outlook on the web, on Windows (new and classic), on Mac, and on mobile devices.<br/><br/>`const myDate = Office.mailbox.item.dateTimeCreated;`<br/>`const myUTCDate = myDate.getUTCDate;`<br/><br/>However, calling  `myDate.getDate` returns a date value in the client computer time zone, which is consistent with the time zone used to display date times values in the Outlook on Windows (classic) and on Mac interfaces, but may be different from the EAC time zone that Outlook on the web, on mobile devices, and in new Outlook on Windows use in its user interface.|If the item is created at 9 AM UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` returns 4am EST.<br/><br/>If the item is modified at 11 AM UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` returns 6am EST.<br/><br/>Notice that if you want to display the creation or modification time in the user interface, you would want to first convert the time to PST to be consistent with the rest of the user interface.|If the item creation time is 9 AM UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` returns 4am EST.<br/><br/>If the item is modified at 11 AM UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` returns 6am EST.|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|Each of the `start` and `end` parameters requires a JavaScript `Date` object. The arguments should be UTC-correct regardless of the time zone used in the user interface of Outlook on the web, on Windows (new and classic), on Mac, or on mobile devices.|If the start and end times for the appointment form are 9 AM UTC and 11 AM UTC, then you should ensure that the `start` and `end` arguments are UTC-correct, which means:<br/><br/><ul><li>`start.getUTCHours` returns 9am UTC</li><li>`end.getUTCHours` returns 11am UTC</li></ul>|If the start and end times for the appointment form are 9 AM UTC and 11 AM UTC, then you should ensure that the `start` and `end` arguments are UTC-correct, which means:<br/><br/><ul><li>`start.getUTCHours` returns 9am UTC</li><li>`end.getUTCHours` returns 11am UTC</li></ul>|

## Helper methods for date-related scenarios

The local time for a user in Outlook on the web, on mobile devices, and in new Outlook on Windows can be different from that in Outlook on Windows (classic) and Outlook on Mac, but the JavaScript **Date** object supports converting to only the client computer time zone or UTC. To overcome this, the Office JavaScript API provides two helper methods: [Office.context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) and [Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

These helper methods take care of any need to handle date or time differently for the following two date-related scenarios in Outlook on the web, on Windows (new and classic), on Mac, and on mobile devices, thus reinforcing "write-once" for different clients of your add-in.

### Scenario A: Displaying item creation or modified time

If you are displaying the item creation time (`Item.dateTimeCreated`) or modification time (`Item.dateTimeModified`in the user interface, first use `convertToLocalClientTime` to convert the `Date` object provided by these properties to obtain a dictionary representation in the appropriate local time. Then display the parts of the dictionary date. The following is an example of this scenario.

```js
// This date is UTC-correct.
const myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In Outlook on Windows (classic) and Outlook on Mac, this dictionary format 
// is in the client computer time zone.
// In Outlook on the web, on mobile devices, or in new Outlook on Windows,
// this dictionary format is in the EAC time zone.
const myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Note that `convertToLocalClientTime` takes care of the difference between Outlook on Windows (classic) or Outlook on Mac and Outlook on the web, on mobile devices, or new Outlook on Windows.

- If `convertToLocalClientTime` detects the current application is Outlook on Windows (classic) or Outlook on Mac, the method converts the `Date` representation to a dictionary representation in the same client computer time zone, consistent with the rest of the Outlook on Windows (classic) or Outlook on Mac user interface.

- If `convertToLocalClientTime` detects the current application is Outlook on the web, on mobile devices, or new Outlook on Windows, the method converts the UTC-correct `Date` representation to a dictionary format in the EAC time zone, consistent with the rest of the user interfaces of these Outlook clients.

### Scenario B: Displaying start and end dates in a new appointment form

If you are obtaining as input different parts of a date-time value represented in the local time, and would like to provide this dictionary input value as a start or end time in an appointment form, first use the `convertToUtcClientTime` helper method to convert the dictionary value to a UTC-correct `Date` object.

In the following example, assume `myLocalDictionaryStartDate` and `myLocalDictionaryEndDate` are date-time values in dictionary format that you have obtained from the user. These values are based on the local time, dependent on the client platform.

```js
const myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
const myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

The resultant values, `myUTCCorrectStartDate` and `myUTCCorrectEndDate`, are UTC-correct. Then, pass these `Date` objects as arguments for the `start` and `end` parameters of the `Mailbox.displayNewAppointmentForm` method to display the new appointment form.

Note that `convertToUtcClientTime` takes care of the difference between Outlook on Windows (classic) or on Mac and Outlook on the web, on mobile devices, or new Outlook on Windows:

- If `convertToUtcClientTime` detects the current application is Outlook on Windows (classic) or on Mac, the method simply converts the dictionary representation to a `Date` object. This `Date` object is UTC-correct, as expected by `displayNewAppointmentForm`.

- If `convertToUtcClientTime` detects the current application is Outlook on the web, on mobile devices, or new Outlook on Windows, the method converts the dictionary format of the date and time values expressed in the EAC time zone to a `Date` object. This `Date` object is UTC-correct, as expected by `displayNewAppointmentForm`.

## See also

- [Deploy and install Outlook add-ins for testing](testing-and-tips.md)
