---
title: Excel JavaScript API requirement set 1.11
description: 'Details about the ExcelApi 1.11 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.11

The ExcelApi 1.11 improved support for comments and workbook-level controls (such as saving and closing the workbook). It also added access to culture settings to help account for localization.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Comment [Mentions](../../excel/excel-add-ins-comments.md#mentions) |Tags and notifies other workbook users through comments. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Comment [Resolution](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Resolve comment threads and get the resolution status. | [Comment](/javascript/api/excel/excel.comment) |
| [Culture settings](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Gets cultural system settings for the workbook, such as number formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Cut and paste (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Replicates the cut-and-paste functionality in Excel for a Range. | [Range](/javascript/api/excel/excel.range) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Save and close workbooks. | [Workbook](/javascript/api/excel/excel.workbook) |
| Worksheet events | Additional events and event information for worksheet calculations and hidden rows. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.11. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.11 or earlier, see [Excel APIs in requirement set 1.11 or earlier](/javascript/api/excel?view=excel-js-1.11&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#excel-excel-application-cultureinfo-member)|Provides information based on current system culture settings.|
||[decimalSeparator](/javascript/api/excel/excel.application#excel-excel-application-decimalseparator-member)|Gets the string used as the decimal separator for numeric values.|
||[thousandsSeparator](/javascript/api/excel/excel.application#excel-excel-application-thousandsseparator-member)|Gets the string used to separate groups of digits to the left of the decimal for numeric values.|
||[useSystemSeparators](/javascript/api/excel/excel.application#excel-excel-application-usesystemseparators-member)|Specifies if the system separators of Excel are enabled.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#excel-excel-comment-mentions-member)|Gets the entities (e.g., people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.comment#excel-excel-comment-resolved-member)|The comment thread status.|
||[richContent](/javascript/api/excel/excel.comment#excel-excel-comment-richcontent-member)|Gets the rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#excel-excel-comment-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Creates a new comment with the given content on the given cell.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-email-member)|The email address of the entity that is mentioned in a comment.|
||[id](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-id-member)|The ID of the entity.|
||[name](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-name-member)|The name of the entity that is mentioned in a comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-mentions-member)|The entities (e.g., people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-resolved-member)|The comment reply status.|
||[richContent](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-richcontent-member)|The rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Creates a comment reply for a comment.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-mentions-member)|An array containing all the entities (e.g., people) mentioned within the comment.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-richcontent-member)|Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an ID attribute of 0, and the second mentioned entity has an ID attribute of 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-name-member)|Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-numberformat-member)|Defines the culturally appropriate format of displaying numbers.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numberdecimalseparator-member)|Gets the string used as the decimal separator for numeric values.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numbergroupseparator-member)|Gets the string used to separate groups of digits to the left of the decimal for numeric values.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))|Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-adjustindent-member(1))|Adjusts the indentation of the range formatting.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-close-member(1))|Close current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-save-member(1))|Save current workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-address-member)|The address of the range that completed calculation.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-changetype-member)|Gets the type of change that represents how the event was triggered.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
