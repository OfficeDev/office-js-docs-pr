---
title: Word JavaScript API requirement set 1.2
description: 'Details about the WordApi 1.2 requirement set'
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
---

# What's new in Word JavaScript API 1.2

WordApi 1.2 added support for inline pictures.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.2. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.2 or earlier, see [Word APIs in requirement set 1.2 or earlier](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Inserts a picture into the body at the specified location.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Inserts an inline picture into the content control at the specified location.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete__)|Deletes the inline picture from the document.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertBreak_breakType__insertLocation_)|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertFileFromBase64_base64File__insertLocation_)|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertHtml_html__insertLocation_)|Inserts HTML at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Inserts an inline picture at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertOoxml_ooxml__insertLocation_)|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertText_text__insertLocation_)|Inserts text at the specified location.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Gets the parent paragraph that contains the inline image.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#select_selectionMode_)|Selects the inline picture.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Inserts a picture at the specified location.|
||[inlinePictures](/javascript/api/word/word.range#inlinePictures)|Gets the collection of inline picture objects in the range.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
