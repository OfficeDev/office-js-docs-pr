# Word add-ins Javascript reference 

Find API reference for the Javascript API for Word for Word add-ins.

_Applies to: Word 2016 for Windows_

## In this section

These are the main objects for the Word Javascript API. For a complete list of topics on MSDN, see the table of contents.

* [Body](Word Add-ins JavaScript Reference/body.md): Represents the body of a document or a section.
* [ContentControl](Word Add-ins JavaScript Reference/contentControl.md): A ContentControl object is a container for content. It is a bound and
 potentially labeled region in a document that serves as a container for specific types of content. For example, content 
 controls can contain contents such as paragraphs of formatted text and other content controls. You can access a 
 content control through the content control collection of the document, document body, paragraph, range, or on a content control.
* [Document](Word Add-ins JavaScript Reference/document.md): The Document object is the top-level object. A Document object contains one or more 
[sections](Word Add-ins JavaScript Reference/section.md), a body that contains the content of the document, and header/footer information.
* [Font](Word Add-ins JavaScript Reference/font.md): The Font object provides text formatting to a body, content control, paragraph, or range.
* [Image](Word Add-ins JavaScript Reference/inlinePicture.md): Represents an inline picture anchored to a paragraph.
* [Paragraph](Word Add-ins JavaScript Reference/paragraph.md): A Paragraph object represents a single paragraph in a selection, range, or document. 
You can access a paragraph through the paragraphs collection in a selection, range, or document. 
* [Range](Word Add-ins JavaScript Reference/range.md): A Range object represents a contiguous area in a document. You get a Range object when you
 get a selection, insert content into the body, insert content into a content control, insert content into a paragraph, 
 or get a search result. You can define and manipulate a range without changing the selection.
* [Section](Word Add-ins JavaScript Reference/section.md):  A Section object is commonly used to define different header and footers as well as 
different page layout configurations of a document. You can access sections from the Document object. 
* [Selection](Word Add-ins JavaScript Reference/document.md#getselection): The document object gives you access to the user's selection in the document, or the current insertion point, if nothing is selected.

## Additional links

* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)