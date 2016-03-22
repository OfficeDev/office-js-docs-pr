
# Labs.Components.InputComponentInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents an instance of an input component.

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## Properties


|Property|Description|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|The underlying [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) object represented by this class.|

## Methods




### constructor

 `function constructor(component: Components.IInputComponentInstance)`

Creates a new [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) instance.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _component_|The [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) from which to create this class.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

Builds a new [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md). Implements the abstract method defined on the base class.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _createAttemptResult_|The result of a create attempt action.|
