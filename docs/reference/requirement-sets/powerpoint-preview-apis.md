---
title: PowerPoint JavaScript preview APIs
description: 'Details about upcoming PowerPoint JavaScript APIs.'
ms.date: 01/27/2021
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
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Specifies the ID of a Slide Layout to be used for the new slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Specifies the ID of a Slide Master to be used for the new slide.|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|Specifies if the bullets in the paragraph are visible.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|Represents the bullet format of the paragraph.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|Represents the horizontal alignment of the paragraph.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Returns the collection of `SlideMaster` objects that are in the presentation.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Returns a collection of tags attached to the presentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Deletes the shape from the shape collection.|
||[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|Returns the fill formatting of this shape.|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|Specifies the height, in points, of the shape.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Gets the unique ID of the shape.|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|The distance, in points, from the left side of the shape to the left side of the slide.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|Returns the line formatting of this shape.|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|Specifies the name of this shape.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Returns a collection of tags in the shape.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|Returns the text frame object of this shape.|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|The distance, in points, from the top edge of the shape to the top edge of the slide.|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|Returns the type of this shape.|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|Specifies the width, in points, of the shape.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|Specifies the height, in points, of the shape.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|Specifies the distance, in points, from the left side of the shape to the left side of the slide.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|Specifies the distance, in points, from the top edge of the shape to the top edge of the slide.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|Specifies the width, in points, of the shape.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|Adds a geometric shape to the slide.|
||[addLine(connectorType?: PowerPoint.ConnectorType, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|Adds a line to the slide.|
||[addTextBox(text: string, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|Adds a text box to the slide with the provided text as the content.|
||[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Gets the number of shapes in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Gets a shape using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Gets a shape using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Gets a shape using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Gets the loaded child items in this collection.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|Returns the fill type of the shape.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|Represents the bold status of font.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|Represents the italic status of font.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|Represents font name (e.g., "Calibri").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|Represents font size in points (e.g., 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|Type of underline applied to the font.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|Represents the dash style of the line.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|Represents the line style of the shape.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|Specifies if the line formatting of a shape element is visible.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|Represents the weight of the line, in points.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Gets the layout of the slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Returns a collection of shapes in the slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Gets the `SlideMaster` object that represents the slide's default content.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Returns a collection of tags in the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Adds a new slide at the end of the collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Gets the unique ID of the slide layout.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Gets the name of the slide layout.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|Returns a collection of shapes in the slide layout.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Gets the number of layouts in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Gets a layout using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Gets a layout using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Gets a layout using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Gets the loaded child items in this collection.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Gets the unique ID of the Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Gets the collection of layouts provided by the Slide Master for slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Gets the unique name of the Slide Master.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|Returns a collection of shapes in the Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Gets the number of Slide Masters in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Gets a Slide Master using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Gets a Slide Master using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Gets a Slide Master using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Gets the loaded child items in this collection.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Gets the unique ID of the tag.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Gets the value of the tag.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Adds a new tag at the end of the collection.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Deletes the tag with the given `key` in this collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Gets the number of tags in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Gets a tag using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Gets a tag using its zero-based index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Gets a tag using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Gets the loaded child items in this collection.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|The automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|Deletes all the text in the text frame.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|Specifies if the text frame contains text.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|Represents the left margin, in points, of the text frame.|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|Represents the right margin, in points, of the text frame.|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|Represents the vertical alignment of the text frame.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|Determines whether lines break automatically to fit text inside the shape.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|Returns a `TextRange` object for the substring in the given range.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|Represents the paragraph format of the text range.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|Represents the plain text content of the text range.|

## See also

- [PowerPoint JavaScript API Reference Documentation](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API requirement sets](powerpoint-api-requirement-sets.md)