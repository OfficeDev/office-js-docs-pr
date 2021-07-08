---
title: Limits for activation and API usage in Outlook add-ins
description: Be aware of certain activation and API usage guidelines, and implement your add-ins to stay within these limits.
ms.date: 06/11/2021
localization_priority: Normal
---

# Limits for activation and JavaScript API for Outlook add-ins

To provide a satisfactory experience for users of Outlook add-ins, you should be aware of certain activation and API usage guidelines, and implement your add-ins to stay within these limits. These guidelines exist so that an individual add-in cannot require Exchange Server or Outlook to spend an unusually long period of time to process its activation rules or calls to the Office JavaScript API, affecting the overall user experience for Outlook and other add-ins. These limits apply to designing activation rules in the add-in manifest, and using custom properties, roaming settings, recipients, Exchange Web Services (EWS) requests and responses, and asynchronous calls.

> [!NOTE]
> If your add-in runs on an Outlook rich client, you must also verify that the add-in performs within certain run-time resource usage limits.

## Limits on where add-ins activate

To learn more about where add-ins do and do not activate, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.

## Limits for activation rules

Follow these guidelines when designing activation rules for Outlook add-ins:

- Limit the size of the manifest to 256 KB. You cannot install the Outlook add-in for an Exchange mailbox if you exceed that limit.

- Specify up to 15 activation rules for the add-in. You cannot install the add-in if you exceed that limit.

- If you use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the body of the selected item, expect an Outlook rich client to apply the rule against only the first 1 MB of the body and not to the rest of the body over that limit. Your add-in would not be activated if matches exist only after the first MB of the body. If you expect that to be a likely scenario, re-design your conditions for activation.

- If you use regular expressions in `ItemHasKnownEntity` or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules, be aware of the following limits and guidelines that generally apply to any Outlook application, and those described in tables 1, 2 and 3 that differ depending on the application.
  - Specify up to only five regular expressions in activation rules in a add-in. You cannot install a add-in if you exceed that limit.
  - Specify regular expressions such that the results you anticipate are returned by the `getRegExMatches` method call within the first 50 matches.
  - Can specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.

Table 1 lists the limits and describes the differences in the support for regular expressions between an Outlook rich client and Outlook on the web or mobile devices. The support is independent of any specific type of device and item body.

**Table 1. General differences in the support for regular expressions**

|Outlook rich client|Outlook on the web or mobile devices|
|:-----|:-----|
|Uses the C++ regular expression engine provided as part of the Visual Studio standard template library. This engine complies with ECMAScript 5 standards. |Uses regular expression evaluation that is part of JavaScript, is provided by the browser, and supports a superset of ECMAScript 5.|
|Because of the different regex engines, expect a regex that includes a custom character class based on predefined character classes may return different results in an Outlook rich client than in Outlook on the web or mobile devices.<br/><br/>As an example, the regex `[\s\S]{0,100}` matches any number, between 0 and 100, of single characters that is a white space or a non-white-space. This regex returns different results in an Outlook rich client than Outlook on the web and mobile devices.<br/><br/>You should rewrite the regex as `(\s\|\S){0,100}` as a work-around. This workaround regex matches any number, between 0 and 100, of white space or non-white space.<br/><br/>You should test each regex thoroughly on each Outlook client, and if a regex returns different results, rewrite the regex. |You should test each regex thoroughly on each Outlook client, and if a regex returns different results, rewrite the regex.|
|By default, limits the evaluation of all regular expressions for an add-in to 1 second. Exceeding this limit causes reevaluation of up to 3 times. Beyond the reevaluation limit, an Outlook rich client disables the add-in from running for the same mailbox in any of the Outlook clients.<br/><br/>Administrators can override these evaluation limits by using the `OutlookActivationAlertThreshold` and `OutlookActivationManagerRetryLimit` registry keys.|Do not support the same resource monitoring or registry settings as in an Outlook rich client. But add-ins with regular expressions that require excessive amount of evaluation time on an Outlook rich client are disabled for the same mailbox on all the Outlook clients.|

Table 2 lists the limits and describes the differences in the portion of the item body that the each of the Outlook applies a regular expression. Some of these limits depend on the type of device and item body, if the regular expression is applied on the item body.

**Table 2. Limits on the size of the item body evaluated**

||Outlook rich client|Outlook on mobile devices|Outlook on the web|
|:-----|:-----|:-----|:-----|
|**Form factor**|Any supported device|Android smartphones, iPad or iPhone|Any supported device other than Android smartphones, iPad and iPhone|
|**Plain text item body**|Applies the regex on the first 1 MB of the data of the body, but not on the rest of the body over that limit.|Activates the add-in only if the body < 16,000 characters.|Activates the add-in only if the body < 500,000 characters.|
|**HTML item body**|Applies the regex on the first 512 KB of the data of the body, but not on the rest of the body over that limit. (The actual number of characters depends on the encoding which can range from 1 to 4 bytes per character.)|Applies the regex on the first 64,000 characters (including HTML tag characters), but not on the rest of the body over that limit.|Activates the add-in only if the body < 500,000 characters.|

Table 3 lists the limits and describes the differences in the matches that each Outlook client returns after evaluating a regular expression. The support is independent of any specific type of device, but may depend on the type of item body, if the regular expression is applied on the item body.

**Table 3. Limits on the matches returned**

||Outlook rich client|Outlook on the web or mobile devices|
|:-----|:-----|:-----|
|**Order of returned matches**|Assume `getRegExMatches` returns matches for the same regular expression applied on the same item is different in an Outlook rich client than in Outlook on the web or mobile devices.|Assume `getRegExMatches` returns matches in different order in an Outlook rich client than in Outlook on the web or mobile devices.|
|**Plain text item body**|`getRegExMatches` returns any matches that are up to 1,536 (1.5 KB) characters, for a maximum of 50 matches.<br/><br/>**Note**: `getRegExMatches` does not return matches in any specific order in the returned array. In general, assume the order of matches in an Outlook rich client for the same regular expression applied on the same item is different from that in Outlook on the web and mobile devices.|`getRegExMatches` returns any matches that are up to 3,072 (3 KB) characters, for a maximum of 50 matches.|
|**HTML item body**|`getRegExMatches` returns any matches that are up to 3,072 (3 KB) characters, for a maximum of 50 matches.<br/> <br/> **Note**: `getRegExMatches` does not return matches in any specific order in the returned array. In general, assume the order of matches in an Outlook rich client for the same regular expression applied on the same item is different from that in Outlook on the web and mobile devices.|`getRegExMatches` returns any matches that are up to 3,072 (3 KB) characters, for a maximum of 50 matches.|

## Limits for JavaScript API

Aside from the preceding guidelines for activation rules, each Outlook client enforces certain limits in the JavaScript object model, as described in Table 4.

**Table 4. Limits to get or set certain data using the Office JavaScript API**

|Feature|Limit|Related API|Description|
|:-----|:-----|:-----|:-----|
|Custom properties|2,500 characters|[CustomProperties](/javascript/api/outlook/office.CustomProperties) object<br/> <br/>[item.loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method|Limit for all custom properties for an appointment or message item. All the Outlook clients return an error if the total size of all custom properties of an add-in exceeds this limit.|
|Roaming settings|32 KB number of characters|[RoamingSettings](/javascript/api/outlook/office.RoamingSettings) object<br/><br/> [context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md#properties) property|Limit for all roaming settings for the add-in. All the Outlook clients return an error if your settings exceed this limit.|
|Extracting well-known entities|2,000 number of characters|[item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method<br/> <br/>[item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method<br/> <br/>[item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method|Limit for Exchange Server to extract well-known entities on the item body. Exchange Server ignores entities beyond that limit. Note that this limit is independent of whether the add-in uses an `ItemHasKnownEntity` rule.|
|Exchange Web Services|1 MB number of characters|[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method|Limit for a request or response to a `Mailbox.makeEwsRequestAsync` call.|
|Recipients|100 recipients|[item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property<br/> <br/>[item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property<br/> <br/>[item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property<br/> <br/>[item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property<br/> <br/>[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-) method<br/> <br/>[Recipient.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) method<br/> <br/>[Recipient.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-) method|Limit for the recipients specified in each property.|
|Display name|255 characters|[EmailAddressDetails.displayName](/javascript/api/outlook/office.emailaddressdetails#displayname) property<br/><br/> [Recipients](/javascript/api/outlook/office.Recipients) object<br/><br/> `item.requiredAttendees` property<br/><br/> `item.optionalAttendees` property <br/><br/>`item.to` property <br/><br/>`item.cc` property|Limit for the length of a display name in an appointment or message.|
|Setting the subject|255 characters|[mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method<br/><br/> [Subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-) method|Limit for the subject in the new appointment form, or for setting the subject of an appointment or message.|
|Setting the location|255 characters|[Location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-) method|Limit for setting the location of an appointment or meeting request.|
|Body in a new appointment form|32 KB number of characters|`Mailbox.displayNewAppointmentForm` method|Limit for the body in a new appointment form.|
|Displaying the body of an existing item|32 KB number of characters|[mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method<br/><br/> [mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method|For Outlook on the web and mobile devices: limit for the body in an existing appointment or message form.|
|Setting the body|1 MB number of characters|[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-) method<br/> <br/>[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)<br/><br/>[Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-) method|Limit for setting the body of an appointment or message item.|
|Number of attachments|499 files on Outlook on the web and mobile devices|[item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method|Limit for the number of files that can be attached to an item for sending. Outlook on the web and mobile devices generally limit attaching up to 499 files, through the user interface and `addFileAttachmentAsync`. An Outlook rich client does not specifically limit the number of file attachments. However, all Outlook clients observe the limit for the size of attachments that user's Exchange Server has been configured with. See the next row for "Size of attachments".|
|Size of attachments|Dependent on Exchange Server|`item.addFileAttachmentAsync` method|There is a limit on the size of all the attachments for an item, which an administrator can configure on the Exchange Server of the user's mailbox.For an Outlook rich client, this limits the number of attachments for an item. For Outlook on the web and mobile devices, the lesser of the two limits - the number of attachments and the size of all attachments - restricts the actual attachments for an item.|
|Attachment filename|255 characters|`item.addFileAttachmentAsync` method|Limit for the length of the filename of an attachment to be added to an item.|
|Attachment URI|2048 characters|`item.addFileAttachmentAsync` method|Limit of the URI of the filename to be added as an attachment to an item.|
|Attachment ID|100 characters|[item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method<br/><br/> [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method|Limit for the length of the ID of the attachment to be added to or removed from an item.|
|Asynchronous calls|3 calls|`item.addFileAttachmentAsync` method<br/><br/>`item.addItemAttachmentAsync` method<br/><br/><br/>`item.removeAttachmentAsync` method<br/><br/> [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-) method<br/><br/>`Body.prependAsync` method<br/><br/>`Body.setSelectedDataAsync` method<br/><br/> [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) method<br/><br/><br/> [item.LoadCustomPropertiesAysnc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method<br/><br/><br/> [Location.getAsync](/javascript/api/outlook/office.Location#getasync-options--callback-) method<br/><br/>`Location.setAsync` method<br/><br/> [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method<br/><br/> [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method<br/><br/> [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method<br/><br/>`Recipients.addAsync` method<br/><br/> [Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) method<br/><br/>`Recipients.setAsync` method<br/><br/> [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) method<br/><br/> [Subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) method<br/><br/>`Subject.setAsync` method<br/><br/> [Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) method<br/><br/> [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-) method|For Outlook on the web or mobile devices: limit of the number of simultaneous asynchronous calls at any one time, as browsers allow only a limited number of asynchronous calls to servers. |

## See also

- [Deploy and install Outlook add-ins for testing](testing-and-tips.md)
- [Privacy, permissions, and security for Outlook add-ins](../concepts/privacy-and-security.md)
