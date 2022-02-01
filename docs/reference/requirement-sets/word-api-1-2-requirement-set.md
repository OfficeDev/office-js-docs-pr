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
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertinlinepicturefrombase64-member(1))|Inserts a picture into the body at the specified location.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertinlinepicturefrombase64-member(1))|Inserts an inline picture into the content control at the specified location.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-delete-member(1))|Deletes the inline picture from the document.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertfilefrombase64-member(1))|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserthtml-member(1))|Inserts HTML at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1))|Inserts an inline picture at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertooxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserttext-member(1))|Inserts text at the specified location.|
||[paragraph](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-paragraph-member)|Gets the parent paragraph that contains the inline image.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-select-member(1))|Selects the inline picture.|
|[Range](/javascript/api/word/word.range)|[inlinePictures](/javascript/api/word/word.range#word-word-range-inlinepictures-member)|Gets the collection of inline picture objects in the range.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertinlinepicturefrombase64-member(1))|Inserts a picture at the specified location.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
