# List Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains a collection of [paragraph](paragraph.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|int|Gets the list's id. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|levelExistences|bool|Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[levelTypes](enums.md)|string|Gets all 9 level types in the list. Each type can be 'Bullet', 'Number' or 'Picture'. Read-only. Possible values are: Bullet, Number, Picture.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets paragraphs in the list. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getLevelParagraphs(level: number)](#getlevelparagraphslevel-number)|[ParagraphCollection](paragraphcollection.md)|Gets the paragraphs that occur at the specified level in the list.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getLevelString(level: number)](#getlevelstringlevel-number)|string|Gets the bullet, number or picture at the specified level as a string.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[setLevelAlignment(level: number, alignment: string)](#setlevelalignmentlevel-number-alignment-string)|void|Sets the alignment of the bullet, number or picture at the specified level in the list.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[setLevelBullet(level: number, listBullet: string, charCode: number, fontName: string)](#setlevelbulletlevel-number-listbullet-string-charcode-number-fontname-string)|void|Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[setLevelIndents(level: number, textIndent: float, textIndent: float)](#setlevelindentslevel-number-textindent-float-textindent-float)|void|Sets the two indents of the specified level in the list.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[setLevelNumbering(level: number, listNumbering: string, formatString: object[])](#setlevelnumberinglevel-number-listnumbering-string-formatstring-object)|void|Sets the numbering format at the specified level in the list.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[setLevelStartingNumber(level: number, startingNumber: number)](#setlevelstartingnumberlevel-number-startingnumber-number)|void|Sets the starting number at the specified level in the list. Default value is 1.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getLevelParagraphs(level: number)
Gets the paragraphs that occur at the specified level in the list.

#### Syntax
```js
listObject.getLevelParagraphs(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|

#### Returns
[ParagraphCollection](paragraphcollection.md)

### getLevelString(level: number)
Gets the bullet, number or picture at the specified level as a string.

#### Syntax
```js
listObject.getLevelString(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|

#### Returns
string

### insertParagraph(paragraphText: string, insertLocation: string)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
listObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|string|Required. The value can be 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Paragraph](paragraph.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### setLevelAlignment(level: number, alignment: string)
Sets the alignment of the bullet, number or picture at the specified level in the list.

#### Syntax
```js
listObject.setLevelAlignment(level, alignment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|
|alignment|string|Required. The level alignment that can be 'left', 'centered' or 'right'. Possible values are: `Unknown` Unknown alignment.,`Left` Alignment to the left.,`Centered` Alignment to the center.,`Right` Alignment to the right.,`Justified` Fully justified alignment.|

#### Returns
void

### setLevelBullet(level: number, listBullet: string, charCode: number, fontName: string)
Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.

#### Syntax
```js
listObject.setLevelBullet(level, listBullet, charCode, fontName);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|
|listBullet|string|Required. The bullet.  Possible values are: Custom , Solid, Hollow, Square, Diamonds, Arrow, Checkmark|
|charCode|number|Optional. Optional. The bullet character's code value. Used only if the bullet is 'Custom'.|
|fontName|string|Optional. Optional. The bullet's font name. Used only if the bullet is 'Custom'.|

#### Returns
void

### setLevelIndents(level: number, textIndent: float, textIndent: float)
Sets the two indents of the specified level in the list.

#### Syntax
```js
listObject.setLevelIndents(level, textIndent, textIndent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|
|textIndent|float|Required. The text indent in points. It is the same as paragraph left indent.|
|textIndent|float|Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.|

#### Returns
void

### setLevelNumbering(level: number, listNumbering: string, formatString: object[])
Sets the numbering format at the specified level in the list.

#### Syntax
```js
listObject.setLevelNumbering(level, listNumbering, formatString);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|
|listNumbering|string|Required. The ordinal format.  Possible values are: None, Arabic, UpperRoman, LowerRoman, UpperLetter, LowerLetter|
|formatString|object[]|Optional. Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.|

#### Returns
void

### setLevelStartingNumber(level: number, startingNumber: number)
Sets the starting number at the specified level in the list. Default value is 1.

#### Syntax
```js
listObject.setLevelStartingNumber(level, startingNumber);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Required. The level in the list.|
|startingNumber|number|Required. The number to start with.|

#### Returns
void
