---
title: Word JavaScript API requirement set 1.2
description: 'Details about the WordApi 1.2 requirement set'
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
---

# What's new in Word JavaScript API 1.2

WordApi 1.2 added support for inline pictures.

## API list

The following table lists the APIs added as part of the WordApi 1.2 requirement set.

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|Deletes the inline picture from the document.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before', or 'After'.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Gets the parent paragraph that contains the inline image. Read-only.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|Selects the inline picture. This causes Word to scroll to the selection.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|Gets the collection of inline picture objects in the range. Read-only.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
