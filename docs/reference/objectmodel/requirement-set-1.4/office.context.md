---
title: Office.context - requirement set 1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
---

# context

### [Office](office.md).context

Office.context provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Properties

| Property | Modes | Return type | Minimum<br>requirement set |
|---|---|---|:---:|
| [contentLanguage](#contentlanguage-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [diagnostics](#diagnostics-contextinformation) | Compose<br>Read | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | Compose<br>Read | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [mailbox](office.context.mailbox.md) | Compose<br>Read | [Mailbox](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [platform](#platform-platformtype) | Compose<br>Read | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [requirements](#requirements-requirementsetsupport) | Compose<br>Read | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Compose<br>Read | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Compose<br>Read | [UI](/javascript/api/office/office.ui?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Property details

#### contentLanguage: String

Gets the locale (language) specified by the user for editing the item.

The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.

##### Type

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
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

#### diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)

Gets information about the environment in which the add-in is running.

##### Type

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### displayLanguage: String

Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.

The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.

##### Type

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
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

#### host: [HostType](/javascript/api/office/office.hosttype)

Gets the Office application host in which the add-in is running.

##### Type

*   [HostType](/javascript/api/office/office.hosttype)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### platform: [PlatformType](/javascript/api/office/office.platformtype)

Provides the platform on which the add-in is running.

##### Type

*   [PlatformType](/javascript/api/office/office.platformtype)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

Provides a method for determining what requirement sets are supported on the current host and platform.

##### Type

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)

Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.

The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.

##### Type

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Minimum permission level](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

<br>

---
---

#### ui: [UI](/javascript/api/office/office.ui)

Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.

##### Type

*   [UI](/javascript/api/office/office.ui)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|
