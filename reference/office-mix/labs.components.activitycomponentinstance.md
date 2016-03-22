
# Labs.Components.ActivityComponentInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents the current instance of an activity component.

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## Properties


|**Name**|**Description**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|The underlying [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) this class represents|

## Methods




### constructor

 `function constructor(component: Components.IActivityComponentInstance)`

Creates a new instance of the [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) class.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _component_|The  **IActivityComponentInstance** to create this class from this class.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

Builds a new  **ActivityComponentAttempt** instance and implements the abstract method defined on the base class

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _createAttemptResult_|The result of a create attempt action.|
