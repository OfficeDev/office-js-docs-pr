# Word add-ins JavaScript reference 

Find API reference for the JavaScript API for Word for Word add-ins.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## In this section

These are the main objects for the Word JavaScript API.

* [Body](word-add-ins-javascript-reference/body.md): Represents the body of a document or a section.
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md): A container for content. It is a bound and
 potentially labeled region in a document that serves as a container for specific types of content. For example, content 
 controls can contain contents such as paragraphs of formatted text and other content controls. You can access a 
 content control through the content control collection of the document, document body, paragraph, range, or on a content control.
* [Document](word-add-ins-javascript-reference/document.md): The top-level object. A Document object contains one or more 
[sections](word-add-ins-javascript-reference/section.md), a body that contains the content of the document, and header/footer information.
* [Font](word-add-ins-javascript-reference/font.md): Provides text formatting to a body, content control, paragraph, or range.
* [Image](word-add-ins-javascript-reference/inlinepicture.md): Represents an inline picture anchored to a paragraph.
* [Paragraph](word-add-ins-javascript-reference/paragraph.md): Represents a single paragraph in a selection, range, or document. 
You can access a paragraph through the paragraphs collection in a selection, range, or document. 
* [Range](word-add-ins-javascript-reference/range.md): Represents a contiguous area in a document. You get a Range object when you
 get a selection, insert content into the body, insert content into a content control, insert content into a paragraph, 
 or get a search result. You can define and manipulate a range without changing the selection.
* [Section](word-add-ins-javascript-reference/section.md):  Defines different headers and footers as well as the different page layout configurations of a document. You can access sections from the Document object. 
* [Selection](word-add-ins-javascript-reference/document.md#getselection): The Document object gives you access to the user's selection in the document or to the current insertion point if nothing is selected.

## Give us your feedback

Your feedback is important to us. 

* Check out the docs and let us know about any questions and issues you find in them by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. Use [this site](http://officespdev.uservoice.com/) for entering your suggestions and ideas.

## Additional resources

* [Word add-ins](word-add-ins.md)
* [Word add-ins programming guide](word-add-ins-programming-guide.md)
* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
* [Snippet Explorer for Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)