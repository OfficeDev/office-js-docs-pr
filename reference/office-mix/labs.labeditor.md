
# Labs.LabEditor

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The  **LabEditor** object allows you to edit a given lab as well as get and set configuration data associated with the lab.

```
class LabEditor
```


## Methods


### getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Retrieves the current lab configuration.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Callback function that is fired once the configuration has been retrieved.|

### setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Sets a new lab configuration.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _configuration_|The configuration to set.|
| _callback_|Callback function that is fired once the configuration has been set.|

### done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indicates that the user has finished editing the lab.

 **Parameters**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Callback function that is fired once the lab editor has finished.|
