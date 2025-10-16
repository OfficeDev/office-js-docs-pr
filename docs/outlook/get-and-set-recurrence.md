---
title: Get and set the recurrence of appointments
description: This topic shows you how to use the Office JavaScript API to get and set various recurrence properties of an appointment using an Outlook add-in.
ms.date: 06/17/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Get and set the recurrence of appointments

Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder. Use the Office JavaScript API to manage the recurrence patterns of an appointment series in your add-in.

> [!NOTE]
> Support for this feature was introduced in [requirement set 1.7](/javascript/api/requirement-sets/outlook/requirement-set-1.7/outlook-requirement-set-1.7). See [clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Recurrence patterns

The recurrence pattern of an appointment is made up of a [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) (such as daily or weekly recurrence) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (for example, the day of the month the appointment occurs).

:::image type="content" source="../images/outlook-recurrence-dialog.png" alt-text="A sample Appointment Recurrence dialog in Outlook.":::

The following table lists the available recurrence types, their configurable properties, and descriptions of their usage.

|Recurrence type|Valid recurrence properties|Usage|
|---|---|---|
|`daily`|<ul><li>[`interval`][interval link]</li></ul>|<ul><li>An appointment occurs every *interval* days. For example, an appointment occurs every ***two*** days.</li></ul>|
|`weekday`|None|<ul><li>An appointment occurs every weekday.</li></ul>|
|`monthly`|<ul><li>[`interval`][interval link]</li><li>[`dayOfMonth`][dayOfMonth link]</li><li>[`dayOfWeek`][dayOfWeek link]</li><li>[`weekNumber`][weekNumber link]</li></ul>|<ul><li>An appointment occurs on day *dayOfMonth* every *interval* months. For example, an appointment occurs on the ***fifth*** day every ***four*** months.</li><li>An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months. For example, an appointment occurs on the ***third Thursday*** every ***two*** months.</li></ul>|
|`weekly`|<ul><li>[`interval`][interval link]</li><li>[`days`][days link]</li></ul>|<ul><li>An appointment occurs on *days* every *interval* weeks. For example, an appointment occurs on ***Tuesday*** and ***Thursday*** every ***two*** weeks.</li></ul>|
|`yearly`|<ul><li>[`interval`][interval link]</li><li>[`dayOfMonth`][dayOfMonth link]</li><li>[`dayOfWeek`][dayOfWeek link]</li><li>[`weekNumber`][weekNumber link]</li><li>[`month`][month link]</li></ul>|<ul><li>An appointment occurs on day *dayOfMonth* of *month* every *interval* years. For example, an appointment occurs on the ***seventh*** day of ***September*** every ***four*** years.</li><li>An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years. For example, an appointment occurs on the ***first*** ***Thursday*** of ***September*** every ***two*** years.</li></ul>|

> [!TIP]
> You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type. The specified day will start the list of days displayed in the recurrence dialog.

## Access the recurrence pattern

As shown in the following table, how you access the recurrence pattern and what you can do with it depends on:

- Whether you're the appointment organizer or an attendee.
- Whether you're using the add-in in compose or read mode.
- Whether the current appointment is a single occurrence or a series.

|Appointment state|Is recurrence editable?|Is recurrence viewable?|
|---|---|---|
|Appointment organizer - compose series|Yes ([`setAsync`][setAsync link])|Yes ([`getAsync`][getAsync link])|
|Appointment organizer - compose instance|No (`setAsync` returns an error)|Yes ([`getAsync`][getAsync link])|
|Appointment attendee - read series|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|
|Appointment attendee - read instance|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|
|Meeting request - read series|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|
|Meeting request - read instance|No (`setAsync` not available)|Yes ([`item.recurrence`][item.recurrence link])|

## Set recurrence as the organizer

Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series. The [SeriesTime][SeriesTime link] object is used to manage that information.

The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only. In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.

```javascript
const seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

const pattern = {
    seriesTime: seriesTimeObject,
    recurrenceType: Office.MailboxEnums.RecurrenceType.Weekly,
    recurrenceProperties:
    {
        interval: 1,
        days: [Office.MailboxEnums.Days.Tue, Office.MailboxEnums.Days.Thu]
    },
    recurrenceTimeZone: { name: Office.MailboxEnums.RecurrenceTimeZone.PacificStandardTime }
};

Office.context.mailbox.item.recurrence.setAsync(pattern, (asyncResult) => {
    console.log(JSON.stringify(asyncResult));
});
```

## Change recurrence as the organizer

In the following example, the appointment organizer gets the [Recurrence](/javascript/api/outlook/office.recurrence) object of an appointment series, then sets a new recurrence duration. This is done in compose mode.

```javascript
Office.context.mailbox.item.recurrence.getAsync((asyncResult) => {
  const recurrencePattern = asyncResult.value;
  recurrencePattern.seriesTime.setDuration(60);
  Office.context.mailbox.item.recurrence.setAsync(recurrencePattern, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log("Failed to set recurrence.");
      return;
    }

    console.log("Successfully set recurrence.");
  });
});
```

## Get recurrence as the organizer

In the following example, the appointment organizer gets the `Recurrence` object of an appointment to determine whether it's a recurring series. This is done in compose mode.

```javascript
Office.context.mailbox.item.recurrence.getAsync((asyncResult) => {
    const recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting.");
    } else {
        console.log(JSON.stringify(recurrence));
    }
});
```

The following example shows the results of the `getAsync` call that retrieves the recurrence of a series.

> [!NOTE]
> In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property. You should use the [SeriesTime][SeriesTime link] methods to get the recurrence date and time properties.

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

## Get recurrence as an attendee

In the following example, an appointment attendee gets the `Recurrence` object of an appointment or meeting request.

```javascript
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    const recurrence = item.recurrence;

    if (recurrence == null) {
        console.log("Non-recurring meeting.");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

The following example shows the value of the `item.recurrence` property of an appointment series.

> [!NOTE]
> In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property. You should use the [SeriesTime][SeriesTime link] methods to get the recurrence date and time properties.

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

## Get the recurrence details

After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence. For example, get the start and end dates and times of the series by using the [SeriesTime][SeriesTime link] methods on the `recurrence.seriesTime` property.

```javascript
// Get the date and time information of the series.
const seriesTime = recurrence.seriesTime;
const startTime = recurrence.seriesTime.getStartTime();
const endTime = recurrence.seriesTime.getEndTime();
const startDate = recurrence.seriesTime.getStartDate();
const endDate = recurrence.seriesTime.getEndDate();
const duration = recurrence.seriesTime.getDuration();

// Get the series time zone.
const timeZone = recurrence.recurrenceTimeZone;

// Get the recurrence properties.
const recurrenceProperties = recurrence.recurrenceProperties;

// Get the recurrence type.
const recurrenceType = recurrence.recurrenceType;
```

## Identify when the recurrence pattern changes

There may be scenarios where you want your add-in to detect and handle changes to the recurrence pattern of a series. For example, you'd like to update the appointment's location if the series is extended. To implement this, you must create a handler for the [RecurrenceChanged](/javascript/api/office/office.eventtype) event. To add an event handler for the `RecurrenceChanged` event, call [Office.context.mailbox.item.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods). When a change is detected, the event handler receives an argument of type [Office.RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs), which provides the updated recurrence object.

The following example shows how to register an event handler for the `RecurrenceChanged` event.

```javascript
// This sample shows how to register an event handler in Outlook.
Office.onReady(() => {
    // Register an event handler to identify when the recurrence pattern of a series is updated.
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, handleEvent, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }

        console.log("Event handler added for the RecurrenceChanged event.");
    });
});

function handleEvent(event) {
    // Get the updated recurrence object.
    const updatedRecurrence = event.recurrence;

    // Perform operations in response to the updated recurrence pattern.
}
```

## Run sample snippets in Script Lab

To test how to get and set the recurrence of an appointment with an add-in, install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and run the following sample snippets.

- "Get recurrence (Read)"
- "Get and set recurrence (Appointment Organizer)"

To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

:::image type="content" source="../images/outlook-recurrence-script-lab.png" alt-text="The recurrence sample snippet in Script Lab.":::

## See also

- [Get or set the time when composing an appointment in Outlook](get-or-set-the-time-of-an-appointment.md)
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)

[getAsync link]: /javascript/api/outlook/office.recurrence#getAsync_options__callback_
[item.recurrence link]: /javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setAsync_recurrencePattern__options__callback_

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayOfMonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayOfWeek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstDayOfWeek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weekNumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
