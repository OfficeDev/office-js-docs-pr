---
title: Office.context - requirement set 1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
---

# context

### [Office](office.md).context

Office.context provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).


##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Members and methods

| Member | Type |
|--------|------|
| [displayLanguage](#displaylanguage-string) | Member |
| [roamingSettings](#roamingsettings-roamingsettings) | Member |

### Namespaces

[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook.

### Members

#### displayLanguage: String

Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.

The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.

##### Type

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)

Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.

The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.

##### Type

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|
