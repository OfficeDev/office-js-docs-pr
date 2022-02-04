---
title: PowerPoint JavaScript preview APIs
description: 'Details about upcoming PowerPoint JavaScript APIs.'
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# PowerPoint JavaScript preview APIs

New PowerPoint JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Slide management | Adds support for adding slides as well as managing slide layouts and slide masters. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Shapes | Adds support for getting references to the shapes in a slide. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## API list

The following table lists the PowerPoint JavaScript APIs currently in preview. For a complete list of all PowerPoint JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#powerpoint-powerpoint-bulletformat-visible-member)|Specifies if the bullets in the paragraph are visible.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-bulletformat-member)|Represents the bullet format of the paragraph.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-horizontalalignment-member)|Represents the horizontal alignment of the paragraph.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-fill-member)|Returns the fill formatting of this shape.|
||[height](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-height-member)|Specifies the height, in points, of the shape.|
||[left](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-left-member)|The distance, in points, from the left side of the shape to the left side of the slide.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-lineformat-member)|Returns the line formatting of this shape.|
||[name](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-name-member)|Specifies the name of this shape.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-textframe-member)|Returns the text frame object of this shape.|
||[top](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-top-member)|The distance, in points, from the top edge of the shape to the top edge of the slide.|
||[type](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-type-member)|Returns the type of this shape.|
||[width](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-width-member)|Specifies the width, in points, of the shape.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-height-member)|Specifies the height, in points, of the shape.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-left-member)|Specifies the distance, in points, from the left side of the shape to the left side of the slide.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-top-member)|Specifies the distance, in points, from the top edge of the shape to the top edge of the slide.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-width-member)|Specifies the width, in points, of the shape.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgeometricshape-member(1))|Adds a geometric shape to the slide.|
||[addLine(connectorType?: PowerPoint.ConnectorType, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addline-member(1))|Adds a line to the slide.|
||[addTextBox(text: string, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1))|Adds a text box to the slide with the provided text as the content.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-clear-member(1))|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-foregroundcolor-member)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-bold-member)|Represents the bold status of font.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-color-member)|HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-italic-member)|Represents the italic status of font.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-name-member)|Represents font name (e.g., "Calibri").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-size-member)|Represents font size in points (e.g., 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-underline-member)|Type of underline applied to the font.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-color-member)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-dashstyle-member)|Represents the dash style of the line.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-style-member)|Represents the line style of the shape.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-visible-member)|Specifies if the line formatting of a shape element is visible.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-weight-member)|Represents the weight of the line, in points.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-deletetext-member(1))|Deletes all the text in the text frame.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-hastext-member)|Specifies if the text frame contains text.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-textrange-member)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-wordwrap-member)|Determines whether lines break automatically to fit text inside the shape.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-font-member)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-getsubstring-member(1))|Returns a `TextRange` object for the substring in the given range.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-paragraphformat-member)|Represents the paragraph format of the text range.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-text-member)|Represents the plain text content of the text range.|

## See also

- [PowerPoint JavaScript API Reference Documentation](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API requirement sets](powerpoint-api-requirement-sets.md)