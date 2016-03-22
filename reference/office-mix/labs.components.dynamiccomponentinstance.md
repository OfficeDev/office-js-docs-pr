
# Labs.Components.DynamicComponentInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents an instance of a dynamic component.

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## Properties


|Property|Description|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|The component instance definition.|

## Methods




### constructor

 `function constructor(component: Components.IDynamicComponentInstance)`

Creates a new dynamic component instance using the [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md) definition.


### getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

Retrieves all of the components created by this dynamic component.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _callback_|Callback function that fires once all of the components have been retrieved.|

### createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

Creates a new component using the dynamic component as component base.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _component_|The component ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)) from which to create the instance.|
| _callback_|Callback function that fires once the component is created.|

### close

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

Indicates there will be no additional submissions associated with this component instance.

 **Parameters**


|Parameter|Description|
|:-----|:-----|
| _callback_|Callback function that fires once the instance is closed.|

### isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

Returns whether the dynamic component is closed. Returns  **true** if closed.

