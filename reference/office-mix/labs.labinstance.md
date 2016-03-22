
# Labs.LabInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

An instance of a lab that is configured for the current user. Use this object to record and retrieve lab data for the user.

```
class LabInstance
```


## Variables


|||
|:-----|:-----|
| `public var data: any`|Container variable for holding user data.|
| `public var components: Labs.ComponentInstanceBase[]`|Components that make up the lab instance.|

## Methods




### getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

Retrieves the current state of the lab for a given user.

 **Parameters**


|||
|:-----|:-----|
| _callback_|The callback function that fires when the lab state is retrieved.|

### setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

Sets the state of the lab for a given user.

 **Parameters**


|||
|:-----|:-----|
| _state_|State to set.|
| _callback_|Callback function that fires once the state is set.|

### Done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indicator function indicating that the user has finished taking the lab.

 **Parameters**


|||
|:-----|:-----|
| _callback_|Callback function that fires once the lab has finished.|
