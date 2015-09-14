*Applies to:* Word 2016

# Word add-ins JavaScript reference 

You'll always start with the ClientRequestContext object. 

```js
    var clientRequestContext = new Word.RequestContext();
```

The **ClientRequestContext** object gives you access to the [Document](Word Add-ins JavaScript Reference/document.md) object which contains the properties and methods that you'll need to extend Word and get access to the following objects and functionality:

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

