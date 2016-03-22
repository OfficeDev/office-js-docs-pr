
# Labs.Timeline

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Provides access to the labs.js timeline feature.

```
class Timeline
```


## Methods




### method

 `function constructor(labsInternal: Labs.LabsInternal)`

Creates a new instance of the  **Timeline** class.


### next

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

Indicates that the timeline should advance to the next slide.

 **Parameters**


|||
|:-----|:-----|
| _completionStatus_|Indicates the current status of the lab.|
| _callback_|Callback function that fires when the lab has moved to the next slide.|
