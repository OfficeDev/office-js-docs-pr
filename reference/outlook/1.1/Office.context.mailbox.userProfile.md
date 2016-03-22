 

# userProfile

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Members

####  displayName :String

Gets the user's display name.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

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
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

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
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```