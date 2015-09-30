# Word add-ins Javascript reference 

Find API reference for the Javascript API for Word for Word add-ins.

_Applies to: Word 2016 for Windows_

## In this section

These are the main objects for the Word Javascript API. For a complete list of topics on MSDN, see the table of contents.

* [Body](word-add-ins-javascript-reference/body.md): Represents the body of a document or a section.
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md): A ContentControl object is a container for content. It is a bound and
 potentially labeled region in a document that serves as a container for specific types of content. For example, content 
 controls can contain contents such as paragraphs of formatted text and other content controls. You can access a 
 content control through the content control collection of the document, document body, paragraph, range, or on a content control.
* [Document](word-add-ins-javascript-reference/document.md): The Document object is the top-level object. A Document object contains one or more 
[sections](word-add-ins-javascript-reference/section.md), a body that contains the content of the document, and header/footer information.
* [Font](word-add-ins-javascript-reference/font.md): The Font object provides text formatting to a body, content control, paragraph, or range.
* [Image](word-add-ins-javascript-reference/inlinepicture.md): Represents an inline picture anchored to a paragraph.
* [Paragraph](word-add-ins-javascript-reference/paragraph.md): A Paragraph object represents a single paragraph in a selection, range, or document. 
You can access a paragraph through the paragraphs collection in a selection, range, or document. 
* [Range](word-add-ins-javascript-reference/range.md): A Range object represents a contiguous area in a document. You get a Range object when you
 get a selection, insert content into the body, insert content into a content control, insert content into a paragraph, 
 or get a search result. You can define and manipulate a range without changing the selection.
* [Section](word-add-ins-javascript-reference/section.md):  A Section object is commonly used to define different headers and footers as well as the different page layout configurations of a document. You can access sections from the Document object. 
* [Selection](word-add-ins-javascript-reference/document.md#getselection): The document object gives you access to the user's selection in the document or to the current insertion point if nothing is selected.

## Additional links

* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)