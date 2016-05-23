
# Work with OneNote page content

In the OneNote add-ins JavaScript API, page content is represented by the following object model.

  ![OneNote page object model diagram](OneNoteOM-page.png)

- A Page object contains a collection of PageContent objects.
- A PageContent object contains a content type of Outline, Image, or Other.
- An Outline object contains a collection of Paragraph objects.
- A Paragraph object contains a content type of RichText, Image, or Other.

To create an empty OneNote page, use one of the following methods:

- [Section.addPage](../../reference/onenote/section#addpagetitle-string)
- [Page.insertPageAsSibling](../../reference/onenote/page#insertpageassiblinglocation-string-title-string)

Then use the following methods to work with the page content. The content and structure of a OneNote page are represented by HTML. Only a [subset of HTML is supported](#supported-html) for creating or updating page content.

- [Page.addOutline](../../reference/onenote/page#addoutlineleft-double-top-double-html-string)
- [Outline.append](../../reference/onenote/outline#appendhtml-string)
- [Outline.prepend](../../reference/onenote/outline#prependhtml-string)
- [Paragraph.insertAsSibling](../../reference/onenote/paragraph#insertassiblinghtml-string-insertlocation-string)
- [Paragraph.delete](../../reference/onenote/paragraph#delete)

## Supported HTML

The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`


//TODO: where get html? update image link
