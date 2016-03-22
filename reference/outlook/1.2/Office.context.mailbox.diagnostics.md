 

# diagnostics

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Provides diagnostic information to an Outlook add-in.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Members

####  hostName :String

Gets a string that represents the name of the host application.

A string that can be one of the following values: `Outlook`, `Mac Outlook` or `OutlookWebApp`.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|
####  hostVersion :String

Gets a string that represents the version of either the host application or the Exchange Server.

If the mail add-in is running on the Outlook desktop client, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|
####  OWAView :String

Gets a string that represents the current view of Outlook Web App.

The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.

If the host application is not Outlook Web App, then accessing this property results in `undefined`.

Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:

*   `OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.
*   `TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.
*   `ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|
