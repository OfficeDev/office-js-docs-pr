---
title: Office.context - preview requirement set
description: ''
ms.date: 11/19/2019
localization_priority: Normal
---

# context

### [Office](Office.md).context

The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Properties

| Property | Return type |
|--------|------|
| [contentLanguage](#contentlanguage-string) | String |
| [diagnostics](#diagnostics-contextinformation) | [ContextInformation](/javascript/api/office/office.contextinformation) |
| [displayLanguage](#displaylanguage-string) | String |
| [host](#host-hosttype) | [HostType](/javascript/api/office/office.hosttype) |
| [officeTheme](#officetheme-officetheme) | [OfficeTheme](/javascript/api/office/office.officetheme) |
| [platform](#platform-platformtype) | [PlatformType](/javascript/api/office/office.platformtype) |
| [requirements](#requirements-requirementsetsupport) | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport) |
| [roamingSettings](#roamingsettings-roamingsettings) | [RoamingSettings](/javascript/api/outlook/office.roamingsettings) |
| [ui](#ui-ui) | [UI](/javascript/api/office/office.ui) |

### Namespaces

[auth](/javascript/api/office/office.auth): Provides support for [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).

[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.

### Property details

#### contentLanguage: String

Gets the locale (language) specified by the user for editing the item.

The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.

##### Type

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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

#### host: [HostType](/javascript/api/office/office.hosttype)

Gets the Office application host in which the add-in is running.

##### Type

*   [HostType](/javascript/api/office/office.hosttype)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)

Provides access to the properties for Office theme colors.

> [!NOTE]
> This member is only supported in Outlook on Windows.

Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.

##### Type

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### Properties:

|Name| Type| Description|
|---|---|---|
|`bodyBackgroundColor`| String|Gets the Office theme body background color as a hexadecimal color triplet.|
|`bodyForegroundColor`| String|Gets the Office theme body foreground color as a hexadecimal color triplet.|
|`controlBackgroundColor`| String|Gets the Office theme control background color as a hexadecimal color triplet.|
|`controlForegroundColor`| String|Gets the Office theme body control color as a hexadecimal color triplet.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| Preview|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
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
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Example

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
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
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|
