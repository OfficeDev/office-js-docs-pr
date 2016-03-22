
# Labs.Components.ActivityComponentAttempt

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents an attempt at completing an activity component.

```
class Permissions
```


## Methods




### constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Creates a new instance of the  **ActivityComponentAttempt** class.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _labs_|Lab instances ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) associated with the component.|
| _componentId_|ID of the component associated with the attempt.|
| _attemptId_|ID of the attempt.|
| _values_|Values, if any, associated with the component.|

### complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

Indicator that the activity has been completed.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Callback function that is invoked once the activity has completed.|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Function that runs over the actions that are retrieved for a given attempt, then populates the state of the lab.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _action_|The action instance ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)).|
