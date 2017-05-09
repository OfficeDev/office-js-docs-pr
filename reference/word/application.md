# Application Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

The Application object.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[createDocument(base64File: string)](#createdocumentbase64file-string)|[Document](document.md)|Creates a new document by using a base64 encoded .docx file.|[1.4](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### createDocument(base64File: string)
Creates a new document by using a base64 encoded .docx file.

#### Syntax
```js
applicationObject.createDocument(base64File);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64File|string|Optional. Optional. The base64 encoded .docx file. The default value is null.|

#### Returns
[Document](document.md)
