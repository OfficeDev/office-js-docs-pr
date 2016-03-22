
# LabsJS.Labs.Components
Provides a high-level overview of the Labs.JS Labs.Components JavaScript API.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

APIs in the Labs.Components module represent the four default components that are presently available for the development of labs (Activity, Choice, Input, and Dynamic components).

## Labs.Components module

Following are Labs.Components types:


### Classes


|||
|:-----|:-----|
|[Labs.Components.ComponentAttempt](../../reference/office-mix/labs.components.componentattempt.md)|Base class for attempts at components.|
|[Labs.Components.ActivityComponentAttempt](../../reference/office-mix/labs.components.activitycomponentattempt.md)|Represents an attempt at completing an activity component.|
|[Labs.Components.ActivityComponentInstance](../../reference/office-mix/labs.components.activitycomponentinstance.md)|Represents the current instance of an activity component.|
|[Labs.Components.ChoiceComponentAnswer](../../reference/office-mix/labs.components.choicecomponentanswer.md)|The answer to a problem presented in a choice component.|
|[Labs.Components.ChoiceComponentAttempt](../../reference/office-mix/labs.components.choicecomponentattempt.md)|Represents an attempt at a choice component.|
|[Labs.Components.ChoiceComponentInstance](../../reference/office-mix/labs.components.choicecomponentinstance.md)|Represents an instance of a choice component.|
|[Labs.Components.ChoiceComponentResult](../../reference/office-mix/labs.components.choicecomponentresult.md)|The result of a choice component submission.|
|[Labs.Components.ChoiceComponentSubmission](../../reference/office-mix/labs.components.choicecomponentsubmission.md)|Represents the submission associated with a choice component.|
|[Labs.Components.DynamicComponentInstance](../../reference/office-mix/labs.components.dynamiccomponentinstance.md)|Represents an instance of a dynamic component.|
|[Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)|Represents the answer to an input component problem.|
|[Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md)|Represents an attempt at interacting with an input component.|
|[Labs.Components.InputComponentInstance](../../reference/office-mix/labs.components.inputcomponentinstance.md)|Represents an instance of an input component.|
|[Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)|The result of an input component submission.|
|[Labs.Components.InputComponentSubmission](../../reference/office-mix/labs.components.inputcomponentsubmission.md)|Represents a submission to an input component.|

### Interfaces


|||
|:-----|:-----|
|[Labs.Components.IActivityComponent](../../reference/office-mix/labs.components.iactivitycomponent.md)|Represents an activity component. Extends [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md).|
|[Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)|Represents a specific instance of an activity component. Extends [Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md).|
|[Labs.Components.IChoice](../../reference/office-mix/labs.components.ichoice.md)|An available choice for a given problem.|
|[Labs.Components.IChoiceComponent](../../reference/office-mix/labs.components.ichoicecomponent.md)|Enables interactions with a choice component.|
|[Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md)|An instance of a choice component.|
|[Labs.Components.IDynamicComponent](../../reference/office-mix/labs.components.idynamiccomponent.md)|Enables interaction with a dynamic component.|
|[Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md)|An instance of a dynamic component.|
|[Labs.Components.IHint](../../reference/office-mix/labs.components.ihint.md)|Hint for a lab problem.|
|[Labs.Components.IInputComponent](../../reference/office-mix/labs.components.iinputcomponent.md)|Enables interacting with an input component.|
|[Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)|An instance of an input component.|
