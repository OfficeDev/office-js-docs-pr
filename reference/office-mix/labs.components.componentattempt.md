
# Labs.Components.ComponentAttempt

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Base class for attempts at components.

```
class ComponentAttempt
```


## Properties


|**Name**|**Description**|
|:-----|:-----|
| `public var _componentId: string`|ID of the specified component.|
| `public var _id: string`|ID of the associated lab.|
| `public var _labs: Labs.LabsInternal`|The lab ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) object that is used to interact with the underlying [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md).|
| `public var _resumed: boolean`|**True** if the lab has resumed progress on a given attempt.|
| `public var _state: Labs.ProblemState`|Current state of the attempt as provided by the enum [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md).|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|Values associated with the attempt, if any, as contained in the [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)object.|

## Methods




### constructor

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Creates a new instance of the ComponentAttempt class and provides input parameter values.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _labs_|The [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) instance to use with the attempt.|
| _attemptId_|The ID associated with the attempt.|
| _values_|Array of values ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)) associated with the attempt.|

### isResumed

 `public function isResumed(): boolean`

Boolean function indicating whether the lab has resumed.  **True** if the lab has resumed.

 **Parameters**

None.


### resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

Indicates whether the lab has resumed progress on the given attempt and loads existing data as part of this process. An attempt must be resumed before it can be used.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Callback function that is fired once the attempt has resumed.|

### getState

 `public function getState(): Labs.ProblemState`

Retrieves the state of the lab.

 **Parameters**

None.


### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Executes the action associated with the attempt.

 **Parameters**

None.


### getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

Retrieves values associated with the attempt

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _key_|The key associated with the value in the value map.|
