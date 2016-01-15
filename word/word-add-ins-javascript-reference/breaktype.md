# BreakType (JavaScript API for Word) 

Specifies the form of a break. 

_Applies to: Word 2016, Word for iPad, Word for Mac_

The following are the supported break types on the API.

| Value         | Description     |
|:-----------------|:--------|
|column| Column break at the insertion point. |
|line| Line break. |
|lineClearLeft|  Line break. |
|lineClearRight|Line break. |
|next| Section break on next page. |
|page| Page break at the insertion point.|
|sectionContinuous| New section without a corresponding page break.|
|sectionEven| String | Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.|
|sectionOdd| String | Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.|
|textWrapping| String | Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.|

## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 