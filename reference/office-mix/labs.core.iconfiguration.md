
# Labs.Core.IConfiguration

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Lab configuration data structure.

```
interface IConfiguration extends Core.IUserData
```


## Properties


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Version of the application associated with this configuration.|
| `components: Core.IComponent[]`|Components included with the lab.|
| `name: string`|The name of the lab.|
| `timeline: Core.ITimelineConfiguration`|The timeline configuration for the lab.|
| `analytics: Core.IAnalyticsConfiguration`|The analytics configuration for the lab.|
