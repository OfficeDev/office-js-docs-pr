
# LabsJS.Labs.Core.Actions
Provides an overview of the LabJS.Labs.Core.Actions JavaScript API.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

These APIs represent the operations of a lab, indicating the lab's current behaviors. The APIs are useful if you are creating new components or developing connections with a new driver (other than Office Mix).

## LabsJS.Labs.Core.Actions API module

The Actions module contains the following types:


### Interfaces


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|The component to close.|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|The component associated with the attempt.|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|The result of creating an attempt for the given component.|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|Creates a new component.|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|The [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) result of creating a new component.|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|The result of a get value action.|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|The result of submitting an answer for an attempt.|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|Options available for the current attempt's timeout action.|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|Options available to the get value operation.|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|Options associated with a resume attempt.|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|Options available for the submit answer action.|

### Variables


|||
|:-----|:-----|
| `var CloseComponentAction: string`|Closes the component and indicates there will be no future actions against it.|
| `var CreateAttemptAction: string`|Action to create a new attempt.|
| `var CreateComponentAction: string`|Action to create a new component.|
| `var AttemptTimeoutAction: string`|Attempt a timeout action.|
| `var GetValueAction: string`|Action to retrieve a value associated with an attempt.|
| `var ResumeAttemptAction: string`|Resume attempt action. Used to indicate the user is resuming work on a given attempt.|
| `var SubmitAnswerAction: string`|Action to submit an answer for a given attempt.|
