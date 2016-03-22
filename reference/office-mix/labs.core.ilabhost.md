
# Labs.Core.ILabHost

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Provides an abstraction layer for connecting Labs.js to the host.

```
interface ILabHost
```


## Methods


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

Retrieves the versions supported by the lab host.

 **Parameters**

None.


### connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

Initializes a connection with the host.

 **Parameters**


|||
|:-----|:-----|
| _versions_|List of host versions that the client can make use of.|
| _callback_|Callback function that fires when the connection is complete.|

### disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

Terminates communication with the host.

 **Parameters**


|||
|:-----|:-----|
| _completionStatus_|Status of the lab at the time of the disconnection.|
| _callback_|Callback function that fires when the disconnect is complete.|

### on

 `on(handler: (string: any, any: any): void)`

Adds an event handler for dealing with messages coming from the host. The resolved promise will be returned back to the host.

 **Parameters**


|||
|:-----|:-----|
| _handler_|The event handler.|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

Sends a message to the host.

 **Parameters**


|||
|:-----|:-----|
| _type_|The type of message being sent.|
| _options_|Message options.|
| _callback_|Callback function that fires once the message is received.|

### create

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

Creates the lab. Stores the host information and sets aside space for storing the configuration and other elements.

 **Parameters**


|||
|:-----|:-----|
| _options_|Options passed as part of the create operation.|
| _callback_|Callback function that fires once the lab has been created.|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

Retrieves the current lab configuration from the host.

 **Parameters**


|||
|:-----|:-----|
| _callback_|Callback function to retrieve the configuration information.|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

Sets a new lab configuration on the host.

 **Parameters**


|||
|:-----|:-----|
| _configuration_|The lab configuration that is set.|
| _callback_|Callback function that fires once the configuration is set.|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

Retrieves the instance configuration for the lab.

 **Parameters**


|||
|:-----|:-----|
| _callback_|Callback function that fires once the configuration instance has been retrieved.|

### getState

 `getState(callback: Core.ILabCallback<any>)`

Retrieves the current state of the lab for a given user.

 **Parameters**


|||
|:-----|:-----|
| _completionStatus_|Callback function that returns the current lab state.|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

Sets the state of the lab for a given user.

 **Parameters**


|||
|:-----|:-----|
| _state_|The lab state.|
| _callback_|Callback function that fires when state has been set.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

Takes an attempt at an action.

 **Parameters**


|||
|:-----|:-----|
| _type_|Type of action.|
| _options_|Options provided with the action.|
| _callback_|Callback function that returns the final executed action.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

Takes an action that has already been completed.

 **Parameters**


|||
|:-----|:-----|
| _type_|Type of action.|
| _options_|Options provided with the action.|
| _result_|Result of the action.|
| _callback_|Callback function that returns the final executed action.|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

Takes an attempt at an action.

 **Parameters**


|||
|:-----|:-----|
| _type_|Type of get action.|
| _options_|Options provided with the get action.|
| _callback_|Callback function that returns the list of completed actions.|
