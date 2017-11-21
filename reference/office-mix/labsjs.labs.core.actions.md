
# LabsJS.Labs.Core.Actions
Provides an overview of the LabJS.Labs.Core.Actions JavaScript API.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

These APIs represent the operations of a lab, indicating the lab's current behaviors. The APIs are useful if you are creating new components or developing connections with a new driver (other than Office Mix).

## LabsJS.Labs.Core.Actions API module

The Actions module contains the following types:


### Interfaces


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.iclosecomponentoptions)|The component to close.|
|[Labs.Core.Actions.ICreateAttemptOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.icreateattemptoptions)|The component associated with the attempt.|
|[Labs.Core.Actions.ICreateAttemptResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.icreateattemptresult)|The result of creating an attempt for the given component.|
|[Labs.Core.Actions.ICreateComponentOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.icreatecomponentoptions)|Creates a new component.|
|[Labs.Core.Actions.ICreateComponentResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.icreatecomponentresult)|The [Labs.Core.IActionResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.iactionresult) result of creating a new component.|
|[Labs.Core.Actions.IGetValueResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.igetvalueresult)|The result of a get value action.|
|[Labs.Core.Actions.ISubmitAnswerResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.isubmitanswerresult)|The result of submitting an answer for an attempt.|
|[Labs.Core.Actions.IAttemptTimeoutOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.iattempttimeoutoptions)|Options available for the current attempt's timeout action.|
|[Labs.Core.Actions.IGetValueOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.igetvalueoptions)|Options available to the get value operation.|
|[Labs.Core.Actions.IResumeAttemptOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.iresumeattemptoptions)|Options associated with a resume attempt.|
|[Labs.Core.Actions.ISubmitAnswerOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.actions.isubmitansweroptions)|Options available for the submit answer action.|

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
