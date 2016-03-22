
# Labs.ComponentInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents an instance of a component, which is an instantiation of a given component for a user at runtime. The object contains a translated view of the component for a specific run of a lab.

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## Properties

None.


## Methods




### Constructor

 `function constructor()`

Initializes a new instance of the  **ComponentInstance** class.


### createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

Creates a new attempt in the context of a component.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Callback fired when the attempt has been created.|

### getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

Retrieves all attempts associated with the given component.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Callback fired when the attempts have been retrieved.|

### getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

Retrieves the default create attempt options. Can be overridden by derived classes.


### buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

Builds an attempt from the given action. Should be implemented by derived classes.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _createAttemptResult_|The create attempt action for the specified attempt.|
