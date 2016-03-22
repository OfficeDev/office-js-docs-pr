
# Labs.Components.InputComponentAttempt

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents an attempt at interacting with an input component.

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## Methods




### constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Creates a new instance of the  **InputComponentAttempt** class.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _labs_|The labs ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) associated with the attempt.|
| _componentID_|ID of the component associated with the attempt.|
| _attemptId_|ID of the specific attempt.|
| _values_|An array containing the value instances ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)).|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Iterates over the retrieved actions for the specified attempt and populates the state of the lab.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _action_|Action associated with the lab state.|

### getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

Retrieves all of the submissions that have previously been submitted for the specified attempt.


### submit

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

Submits a new answer that was graded by the lab and will not use the host to compute a grade.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _answer_|The answer associated with the attempt.|
| _result_|The result associated with the submission.|
| _callback_|Callback function that fires once the submission has been received.|
