# Word JavaScript APIs

Welcome to the Word JavaScript API documentation repository. Here you'll find what you'll need to create the next generation of Word add-ins in Office 2016 for Windows (if you can't find it, open an issue). The new Word JavaScript APIs provide Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. 

This documentation is [published on MSDN](https://msdn.microsoft.com/EN-US/library/office/mt616496.aspx). 

## Introduction to Word JS APIs 1.3 
This branch contains the new APIs that our team is working on. We are plan to ship these changes in the next few months. This is a great time to give feedback on these APIs.

This section describes the new set of Word JavaScript APIs that are being planned for the next release (Requirement Set 1.3). Please review and provide your feedback. Provide your feedback by opening new issues in GitHub using the links in the table. 

_**Note**: The listed features are still under the design and review phase and are not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._


| New feature	   | Description	| Give feedback|
|:---------------|:--------|:----------|
|[Application object](word-add-ins-javascript-reference/application.md)| Represents the Microsoft Word application. This currently enables you to create a new Word document in memory. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewApplicationObject)_|
|Get child ranges| Gets a range with an origin at either the beginning or end of an object.<br/><br/> [Body.GetChildRange(RangeOrigin rangeOrigin, int length)](word-add-ins-javascript-reference/body.md#getchildrangerangeorigin-string-length-number) <br/><br/>[ContentControl.GetChildRange(RangeOrigin rangeOrigin, int length)](word-add-ins-javascript-reference/contentcontrol.md#getchildrangerangeorigin-string-length-number) <br/><br/> [Paragraph.GetChildRange(RangeOrigin rangeOrigin, int length)](word-add-ins-javascript-reference/paragraph.md#getchildrangerangeorigin-string-length-number) <br/><br/> [Range.GetChildRange(RangeOrigin rangeOrigin, int length)](word-add-ins-javascript-reference/range.md#getchildrangerangeorigin-string-length-number) | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-GetChildRanges)_|
|Get ranges via deliliters | Gets a range collection from an object by using delimiters.<br/><br/>[Body.GetRanges(string[] delimiters, [Optional] bool excludeDelimiters, [Optional] bool trimWhite, [Optional] bool excludeEndingMarks)](word-add-ins-javascript-reference/body.md#getrangesdelimiters-string-excludedelimiters-bool-trimwhite-bool-excludeendingmarks-bool) <br/><br/>  [ContentControl.GetRanges(string[] delimiters, [Optional] bool excludeDelimiters, [Optional] bool trimWhite, [Optional] bool excludeEndingMarks, [Optional] bool within)](word-add-ins-javascript-reference/contentcontrol.md#getrangesdelimiters-string-excludedelimiters-bool-trimwhite-bool-excludeendingmarks-bool) <br/><br/> [Paragraph.GetRanges(string[] delimiters, [Optional] bool excludeDelimiters, [Optional] bool trimWhite)](word-add-ins-javascript-reference/paragraph.md#getrangesdelimiters-string-excludedelimiters-bool-trimwhite-bool) <br/><br/> [Range.GetRanges(string[] delimiters, [Optional] bool excludeDelimiters, [Optional] bool trimWhite, [Optional] bool excludeEndingMarks, [Optional] bool within)](word-add-ins-javascript-reference/range.md#getrangesdelimiters-string-excludedelimiters-bool-trimwhite-bool-excludeendingmarks-bool) | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-GetRangesViaDelimiters)_|
|Range comparison and expansion| Compare whether a range contains another range.<br/> [Range.HasRange(Range range, [Optional] bool isSubset)](word-add-ins-javascript-reference/range.md#hasrangerange-range-issubset-bool)  <br/><br/> Expand the range to include the bounds of another range.<br/> [Range.ExpandTo(Range range)](word-add-ins-javascript-reference/range.md#expandtorange-range) <br/><br/>Adjust the start and end position of a range. <br/> [Range.Adjust(int startAdjust, int endAdjust)](word-add-ins-javascript-reference/range.md#adjuststartadjust-number-endadjust-number)  | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-RangeCompareAndExpansion)_|
|Create and open a new Word Document| This feature enables you to create, change, open, and make changes to a Word document (.docx) before it is displayed in the Word UI. This feature is part of the new [Application](word-add-ins-javascript-reference/application.md) object. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewWordDocument)_|
|Strongly typed List objects| 	This gives you access to ordered and unordered list objects. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewListObject)_|
|Strongly typed Table objects| 	This gives you access to table objects. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewTableObject)_|


## Try it out 

_**Note**: New features cannot be tried out right now._


We've been working on a Snippet Explorer to let you browse through common code snippets and learn how the new APIs work. Give it a try. The code snippets referenced by the Snippet Explorer are available [here](https://officesnippetexplorer.azurewebsites.net/#/snippets/word). 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know about any [issues](https://github.com/OfficeDev/office-js-docs/issues) you find in those docs by submitting issues directly in this repository. Let us know what you think about the APIs and the general programming experience. 

We suggest you use the tags [office-js] and [word] on StackOverflow for asking questions to the community.
