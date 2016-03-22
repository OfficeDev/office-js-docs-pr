
# Labs.Components.InputComponentSubmission

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents a submission to an input component.

```
class InputComponentSubmission
```


## Properties


|Property|Description|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|The answer ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)) associated with the submission.|
| `public var result: Components.InputComponentResult`|The result ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)) of the submission.|
| `public var time: number`|The time at which the submission was received.|

## Methods




### constructor

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

Creates a new instance of the  **InputComponentSubmission** class.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _answer_|The answer associated with the submission.|
| _result_|The result of the submission.|
| _time_|The time at which the submission was received.|
