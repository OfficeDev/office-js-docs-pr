# Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

Word Add-ins run across multiple versions of Office including Office 2016 for Windows, Office for the iPad, Office for the Mac, and Office Online. The following table lists the Word requirement sets, the Office host applications that support that requirement set, and the build or version numbers.

|  Requirement set  |   Office 2016 for Windows  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |
|:-----|-----|:-----|:-----|:-----|
| Word Api 1.3  | In preview, Build 6925 or later| In preview, May 2016, 1.21 or later | In preview, 15.22 or later| We're working on it | 
| Word Api 1.2  | December 2015 update, Build 6568 or later | January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 | 
| Word API 1.1  | Shipped with Office 2016 <br>Version 1509 (Build 16.0.4266.1001)</br> or later| January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 | 

> **Note**: The build number for Office 2016 install via MSI is 16.0.4266.1001. 

To find out more about versions and build numbers, see:
- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Where you can find the version and build number for an Office 365 client application)

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## What's new in Word JavaScript API 1.3 
The following are the new additions to the Word JavaScript APIs in requirement set 1.3. 
Word1.2
|Object| What is new| Description|Req. Set|
|:----|:----|:----|:----|
|[contentControl](reference/word/contentcontrol.md)|_Method_ > [insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](reference/word/contentcontrol.md#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Relationship_ > paragraph|Gets the parent paragraph that contains the inline image. Read-only.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Relationship_ > parentTable|Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.|1.3|
|[inlinePicture](reference/word/inlinepicture.md)|_Relationship_ > parentTableCell|Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.|1.3|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [delete()](reference/word/inlinepicture.md#delete)|Deletes the inline picture from the document.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertBreak(breakType: BreakType, insertLocation: InsertLocation)](reference/word/inlinepicture.md#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](reference/word/inlinepicture.md#insertfilefrombase64base64file-string-insertlocation-insertlocation)|Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertHtml(html: string, insertLocation: InsertLocation)](reference/word/inlinepicture.md#inserthtmlhtml-string-insertlocation-insertlocation)|Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](reference/word/inlinepicture.md#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertOoxml(ooxml: string, insertLocation: InsertLocation)](reference/word/inlinepicture.md#insertooxmlooxml-string-insertlocation-insertlocation)|Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertParagraph(paragraphText: string, insertLocation: InsertLocation)](reference/word/inlinepicture.md#insertparagraphparagraphtext-string-insertlocation-insertlocation)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [insertText(text: string, insertLocation: InsertLocation)](reference/word/inlinepicture.md#inserttexttext-string-insertlocation-insertlocation)|Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](reference/word/inlinepicture.md)|_Method_ > [select(selectionMode: SelectionMode)](reference/word/inlinepicture.md#selectselectionmode-selectionmode)|Selects the inline picture. This causes Word to scroll to the selection.|1.2|
|[range](reference/word/range.md)|_Relationship_ > inlinePictures|Gets the collection of inline picture objects in the range. Read-only.|1.2|
|[range](reference/word/range.md)|_Method_ > [insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](reference/word/range.md#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|1.2|

## What's new in Word JavaScript API 1.2
The following are the new additions to the Word JavaScript APIs in requirement set 1.2. 

## Additional resources

- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
