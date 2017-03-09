# BreakType (JavaScript API for Word)

Specifies the form of a break.

_Applies to: Word 2016, Word for iPad, Word for Mac, Word Online_

The following are the supported break types on the API.

| **Value**         | **Type** | **Description**     |
|:-----------------|:--------|:----|
|line| | Line break. |
|page| | Page break at the insertion point.|
|sectionNext| | Section break on next page. Type of next will be obsoleted.|
|sectionContinuous| | New section without a corresponding page break.|
|sectionEven| string | Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.|
|sectionOdd| string | Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.|
|textWrapping| string | Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.|

## Support details
Use the [requirement set](../office-add-in-requirement-sets.md) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).
