
# Host element
Specifies the type of Office host application your Office Add-in supports.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|string|required|The name of the type of Office host application.|

## Remarks

You can specify the following values in the  **Name** attribute of a **Host** element. Each value maps to the set of one or more Office host applications your add-in supports.



|**Name**|**Office host applications**|
|:-----|:-----|
| `"Document"`|Word, Word Online, Word on iPad|
| `"Database"`|Access web apps|
| `"Mailbox"`|Outlook, Outlook Web App, OWA for Devices|
| `"Presentation"`|PowerPoint, PowerPoint Online, PowerPoint on iPad|
| `"Project"`|Project|
| `"Workbook"`|Excel, Excel Online, Excel on iPad|

## Remarks

For more information about specifying host support, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

