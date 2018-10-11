# Outlook add-in API requirement set 1.7

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

## What's new in 1.7?

Requirement set 1.7 includes all of the features of [Requirement set 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). It added the following features.

- Added new APIs regarding the recurrence pattern on appointments and messages that are meeting requests.
- Modified the item.from property to also be available in Compose mode.
- Added support for RecurrenceChanged, RecipientsChanged, and AppointmentTimeChanged events.

### Change log

- Added [From](/javascript/api/outlook_1_7/office.from): Adds a new object that provides a method to get the from value.
- Added [Organizer](/javascript/api/outlook_1_7/office.organizer): Adds a new object that provides a method to get the organizer value.
- Added [Recurrence](/javascript/api/outlook_1_7/office.recurrence): Adds a new object that provides methods to get and set the recurrence pattern of appointments but only get the recurrence pattern of messages that are meeting requests.
- Added [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): Adds a new object that represents the time zone configuration of the recurrence pattern.
- Added [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): Adds a new object that provides methods to get and set the dates and times of appointments in a recurring series and to get the dates and times of meeting requests in a recurring series.
- Added [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): Adds a new method that adds an event handler for a supported event.
- Modified [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): Modifies to get the from value in Compose mode.
- Modified [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - Modifies to get the organizer value in Compose mode.
- Added [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): Adds a new property that gets or sets an object which provides methods to manage the recurrence pattern of an appointment item. This property can also be used to get the recurrence pattern of a meeting request item.
- Added [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): Adds a new method that removes an event handler.
- Added [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): Adds a new property that gets the id of the series an occurrence belongs to.
- Added [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): Adds a new enum that specifies the day of week or type of day.
- Added [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): Adds a new enum that specifies the month.
- Added [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): Adds a new enum that specifies the time zone applied to the recurrence.
- Added [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): Adds a new enum that specifies the type of recurrence.
- Added [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): Adds a new enum that specifies the week of the month.
- Modified [Office.EventType](/javascript/api/office/office.eventtype): Modifies to support RecurrenceChanged, RecipientsChanged, and AppointmentTimeChanged events through addition of `RecurrenceChanged`, `RecipientsChanged`, and `AppointmentTimeChanged` entries respectively.

## See also

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)