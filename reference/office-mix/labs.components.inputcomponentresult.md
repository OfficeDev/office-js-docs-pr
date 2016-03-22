
# Labs.Components.InputComponentResult

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The result of an input component submission.

```
class InputComponentResult
```


## Properties


|Property|Description|
|:-----|:-----|
| `public var score: any`|The score associated with the submission.|
| `public var complete: boolean`|Indicates whether the result submitted resulted in the completion of the attempt.  **True** if the attempt is completed.|

## Methods




### constructor

 `function constructor(score: any, complete: boolean)`

Creates a new instance of the  **InputComponentResult** class.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _score_|The score associated with the result.|
| _complete_|Boolean  **true** if the result completed the attempt.|
