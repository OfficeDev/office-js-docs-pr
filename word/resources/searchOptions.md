# SearchOptions  
Specifies the options to be included in a search operation.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`ignorePunct`| bool | Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.|  
|`ignoreSpace`| bool |Gets or sets a value that indicates whether to ignore all white space between words. Corresponds to the Ignore white-space characters check box in the Find and Replace dialog box.|
|`matchCase`| bool |Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).| 
|`matchPrefix`| bool  |Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box. |
|`matchSoundsLike`| bool |Gets or sets a value that indicates whether to find words that sound similar to the search string. Corresponds to the Sounds like check box in the Find and Replace dialog box | 
|`matchSuffix`| bool |Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box. | 
|`matchWholeWord`| bool |Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box. | 
|`matchWildcards`| bool |Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box. |  |



## Wildcard Guidance 

| To find:         | Wildcard |  Sample |
|:-----------------|:--------|:----------|
| Any single character| ? |s?t finds sat and set. |
|Any string of characters| * |s*d finds sad and started.|
|The beginning of a word|< |<(inter) finds interesting and intercept, but not splintered.|
|The end of a word |> |(in)> finds in and within, but not interesting.|
|One of the specified characters|[ ] |w[io]n finds win and won.|
|Any single character in this range| [-] |[r-t]ight finds right and sight. Ranges must be in ascending order.|
|Any single character except the characters in the range inside the brackets|[!x-z] |t[!a-m]ck finds tock and tuck, but not tack or tick.|
|Exactly n occurrences of the previous character or expression|{n} |fe{2}d finds feed but not fed.|
|At least n occurrences of the previous character or expression|{n,} |fe{1,}d finds fed and feed.|
|From n to m occurrences of the previous character or expression|{n,m} |10{1,3} finds 10, 100, and 1000.|
|One or more occurrences of the previous character or expression|@ |lo@t finds lot and loot.|
























