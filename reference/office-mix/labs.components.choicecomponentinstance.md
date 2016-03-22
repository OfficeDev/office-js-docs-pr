
# Labs.Components.ChoiceComponentInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents an instance of a choice component.

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## Properties


|Property|Description|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|The underlying [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) which this class represents.|

## Methods




### constructor

 `function constructor(component: Components.IChoiceComponentInstance)`

Creates a new instance of the  **ChoiceComponentInstance** class.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _component_|The [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) object from which to create this class.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

Builds a new  **ChoiceComponentAttempt** instance and implements the abstract method defined on the base class.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _createAttemptResult_|The result from the create attempt action.|
