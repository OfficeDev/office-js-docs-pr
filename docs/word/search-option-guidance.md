---
title: Use search options in your Word add-in to find text 
description: Learn to use search options in your Word add-in.
ms.date: 02/28/2022
ms.localizationpriority: medium
---

# Use search options in your Word add-in to find text

Add-ins frequently need to act based on the text of a document. A search method is exposed by every content control (this includes [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), and the base [ContentControl](/javascript/api/word/word.contentcontrol) object). This method takes in a string (or wildcard expression) representing the text you are searching for and a [SearchOptions](/javascript/api/word/word.searchoptions) object. It returns a collection of ranges which match the search text.

## Search options

The search options are a collection of boolean values defining how the search parameter should be treated.

| Property       | Description|
|:---------------|:----|
|ignorePunct|Gets or sets a value indicating whether to ignore all punctuation characters between words. Corresponds to the "Ignore punctuation characters" checkbox in the **Find and Replace** dialog box.|
|ignoreSpace|Gets or sets a value indicating whether to ignore all whitespace between words. Corresponds to the "Ignore white-space characters" checkbox in the **Find and Replace** dialog box.|
|matchCase|Gets or sets a value indicating whether to perform a case-sensitive search. Corresponds to the "Match case" checkbox in the **Find and Replace** dialog box.|
|matchPrefix|Gets or sets a value indicating whether to match words that begin with the search string. Corresponds to the "Match prefix" checkbox in the **Find and Replace** dialog box.|
|matchSuffix|Gets or sets a value indicating whether to match words that end with the search string. Corresponds to the "Match suffix" checkbox in the **Find and Replace** dialog box.|
|matchWholeWord|Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word. Corresponds to the "Find whole words only" checkbox in the **Find and Replace** dialog box.|
|matchWildcards|Gets or sets a value indicating whether the search will be performed using special search operators. Corresponds to the "Use wildcards" checkbox in the **Find and Replace** dialog box.|

## Wildcard guidance

The following table provides guidance around the Word JavaScript API's search wildcards.

| To find:         | Wildcard |  Sample |
|:-----------------|:--------|:----------|
|Any single character| ? |s?t finds sat and set. |
|Any string of characters| * |s*d finds sad and started.|
|The beginning of a word|< |<(inter) finds interesting and intercept, but not splintered.|
|The end of a word |> |(in)> finds in and within, but not interesting.|
|One of the specified characters|[ ] |w[io]n finds win and won.|
|Any single character in this range| [-] |[r-t]ight finds right and sight. Ranges must be in ascending order.|
|Any single character except the characters in the range inside the brackets|[!x-z] |t[!a-m]ck finds tock and tuck, but not tack or tick.|
|Exactly *n* occurrences of the previous character or expression|{n} |fe{2}d finds feed but not fed.|
|At least *n* occurrences of the previous character or expression|{n,} |fe{1,}d finds fed and feed.|
|From *n* to *m* occurrences of the previous character or expression|{n,m} |10{1,3} finds 10, 100, and 1000.|
|One or more occurrences of the previous character or expression|@ |lo@t finds lot and loot.|

### Escaping special characters

Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a backslash character ('\\'), but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.

## Examples

The following examples demonstrate common scenarios.

### Ignore punctuation search

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### Search based on a prefix

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### Search based on a suffix

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### Search using a wildcard

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

More information can be found in the [Word JavaScript Reference API](../reference/overview/word-add-ins-reference-overview.md).
