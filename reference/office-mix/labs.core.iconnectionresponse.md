
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
| `hostVersion: Core.IVersion`|Version information ([Labs.Core.IVersion](https://dev.office.com/reference/add-ins/office-mix/labs.core.iversion)) for the server.|
| `userInfo: Core.IUserInfo`|Information about the user ([Labs.Core.IUserInfo](https://dev.office.com/reference/add-ins/office-mix/labs.core.iuserinfo)).|
