
# userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Members and methods

| Member | Type |
|--------|------|
| [accountType](#accounttype-string) | Member |
| [displayName](#displayname-string) | Member |
| [emailAddress](#emailaddress-string) | Member |
| [timeZone](#timezone-string) | Member |

### Members

####  accountType :String

> [!NOTE]
> This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).

Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.

| Value | Description |
|-------|-------------|
| `enterprise` | The mailbox is on an on-premises Exchange server. |
| `gmail` | The mailbox is associated with a Gmail account. |
| `office365` | The mailbox is associated with an Office 365 work or school account. |
| `outlookCom` | The mailbox is associated with a personal Outlook.com account. |

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  displayName :String

Gets the user's display name.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  emailAddress :String

Gets the user's SMTP email address.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  timeZone :String

Gets the user's default time zone.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```