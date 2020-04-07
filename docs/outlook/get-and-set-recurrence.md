---
title: Get and set recurrence in an Outlook add-in
description: This topic shows you how to use the Office JavaScript API to get and set various recurrence properties of an item in an Outlook add-in.
ms.date: 01/14/2020
localization_priority: Normal
---

# Get and set recurrence

Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder. You can use the Office JavaScript API to manage the recurrence patterns of an appointment series in your add-in.

> [!NOTE]
> Support for this feature was introduced in requirement set 1.7. See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Available recurrence patterns

To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).

**Table 1. Recurrence types and their applicable properties**

|Recurrence type|Valid recurrence properties|Usage|
|---|---|---|
|`daily`|- [`interval`][interval link]|An appointment occurs every *interval* days. Example: An appointment occurs every **_2_** days.|
|`weekday`|None.|An appointment occurs every weekday.|
|`monthly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]|- An appointment occurs on day *dayOfMonth* every *interval* months. Example: An appointment occurs on day **_5_** every **_4_** months.<br/><br/>- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months. Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.|
|`weekly`|- [`interval`][interval link]<br/>- [`days`][days link]|An appointment occurs on *days* every *interval* weeks. Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.|
|`yearly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]<br/>- [`month`][month link]|- An appointment occurs on day *dayOfMonth* of *month* every *interval* years. Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.<br/><br/>- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years. Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.|

> [!NOTE]
> You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type. The specified day will start the list of days displayed in the recurrence dialog.

## Access recurrence

How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.

**Table 2. Applicable appointment states**

|Appointment state|Is recurrence editable?|Is recurrence viewable?|
|---|---|---|
|Appointment organizer - compose series|Yes ([`setAsync`][setAsync link])|Yes ([`getAsync`][getAsync link])|
|Appointment organizer - compose instance|No (`setAsync` returns an error)|Yes ([`getAsync`][getAsync link])|
|Appointment attendee - read series|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|
|Appointment attendee - read instance|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|
|Meeting request - read series|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|
|Meeting request - read instance|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|

## Set recurrence as the organizer

Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series. The [`SeriesTime`][SeriesTime link] object is used to manage that information.

The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only. In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.

```js
var seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

var pattern = {
    "seriesTime": seriesTimeObject,
    "recurrenceType": "weekly",
    "recurrenceProperties": {"interval": 1, "days": ["tue", "thu"]},
    "recurrenceTimeZone": {"name": "Pacific Standard Time"}};

Office.context.mailbox.item.recurrence.setAsync(pattern, callback);

function callback(asyncResult)
{
    console.log(JSON.stringify(asyncResult));
}
```

## Get recurrence

### Get recurrence as the organizer

In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    var context = asyncResult.context;
    var recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.

> [!NOTE]
> In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property. You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### Get recurrence as an attendee

In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    var recurrence = item.recurrence;
    var seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

The following example shows the value of the `item.recurrence` property for an appointment series.

> [!NOTE]
> In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property. You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### Get the recurrence details

After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence. For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.

```js
// Get series date and time info
var seriesTime = recurrence.seriesTime;
var startTime = recurrence.seriesTime.getStartTime();
var endTime = recurrence.seriesTime.getEndTime();
var startDate = recurrence.seriesTime.getStartDate();
var endDate = recurrence.seriesTime.getEndDate();
var duration = recurrence.seriesTime.getDuration();

// Get series time zone
var timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
var recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
var recurrenceType = recurrence.recurrenceType;
```

## See also

[RecurrenceChanged event](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getasync-options--callback-
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setasync-recurrencepattern--options--callback-

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayofmonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayofweek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstdayofweek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weeknumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
