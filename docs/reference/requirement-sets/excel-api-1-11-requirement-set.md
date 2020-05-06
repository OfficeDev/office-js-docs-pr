---
title: Excel JavaScript API requirement set 1.11
description: 'Details about the ExcelApi 1.11 requirement set'
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.11

The ExcelApi 1.11 improved support for comments and workbook-level controls (such as saving and closing the workbook). It also added access to culture settings to help account for localization.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Comment [Mentions](../../excel/excel-add-ins-comments.md#mentions) |Tags and notifies other workbook users through comments. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Comment [Resolution](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Resolve comment threads and get the resolution status. | [Comment](/javascript/api/excel/excel.comment) |
| [Culture settings](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Gets cultural system settings for the workbook, such as number formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Cut and paste (moveTo)](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | Replicates the cut-and-paste functionality in Excel for a Range. | [Range](/javascript/api/excel/excel.range) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Save and close workbooks. | [Workbook](/javascript/api/excel/excel.workbook) |
| Worksheet events | Additional events and event information for worksheet calculations and hidden rows. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.11. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.11 or earlier, see [Excel APIs in requirement set 1.10 or earlier](/javascript/api/excel?view=excel-js-1.11).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Gets the string used as the decimal separator for numeric values. This is based on Excel's local settings.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on Excel's local settings.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Specifies if the system separators of Excel are enabled.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Gets the entities (e.g., people) that are mentioned in comments.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Gets the rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[resolved](/javascript/api/excel/excel.comment#resolved)|The comment thread status. A value of "true" means the comment thread is resolved.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Creates a new comment with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|The email address of the entity that is mentioned in comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|The id of the entity. The id matches one of the ids in `CommentRichContent.richContent`.|
||[name](/javascript/api/excel/excel.commentmention#name)|The name of the entity that is mentioned in comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|The entities (e.g., people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentreply#resolved)|The comment reply status. A value of "true" means the reply is in the resolved state.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|The rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Creates a comment reply for comment.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|An array containing all the entities (e.g., people) mentioned within the comment.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an id attribute of 0, and the second mentioned entity has an id attribute of 1.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us"). This is based on current system settings.|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Gets the string used as the decimal separator for numeric values. This is based on current system settings.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on current system settings.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250 and is measured in characters.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|The address of the range that completed calculation.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Gets the type of change that represents how the event was triggered. See `Excel.RowHiddenChangeType` for details.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.11)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)