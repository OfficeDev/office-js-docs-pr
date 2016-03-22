
# Labs.Core.IConfigurationInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Base class for instances of a lab configuration. An instance is an instantiation of a configuration for a given user and contains a translated view of the configuration for a particular run of the lab. This view may exclude hidden information (for example, hints and answers) and also contains IDs to identify the various instances.

```
interface IConfigurationInstance extends Core.IUserData
```


## Properties


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Version of the lab associated with this configuration.|
| `components: Core.IComponentInstance[]`|Components associated with the lab.|
| `name: string`|Name of the lab.|
| `timeline: Core.ITimelineConfiguration`|Timeline configuration for the lab.|
