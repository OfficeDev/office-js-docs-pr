
# Labs.Core.IConnectionResponse

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Response information returning from a connection call.

```
interface IConnectionResponse
```


## Properties


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|Initialization configureation information, or  **null** if the app has not been initialized.|
| `mode: Core.LabMode`|The mode which the lab is currently running in.|
| `hostVersion: Core.IVersion`|Version information ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)) for the server.|
| `userInfo: Core.IUserInfo`|Information about the user ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)).|
