
# Labs.Core.ITimelineConfiguration

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Configuration options for the [Labs.Timeline](https://dev.office.com/reference/add-ins/office-mix/labs.timeline). Allows you to specify a set of timeline configuration options.

```
interface ITimelineConfiguration
```


## Properties


|||
|:-----|:-----|
| `duration: number`|The duration of the lab, in seconds.|
| `capabilities: string[]`|An array list of timeline capabilities that the lab supports, for example, play, pause, seek, and so forth.|
