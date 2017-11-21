
# LabsJS.Labs.Components
Provides a high-level overview of the Labs.JS Labs.Components JavaScript API.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

APIs in the Labs.Components module represent the four default components that are presently available for the development of labs (Activity, Choice, Input, and Dynamic components).

## Labs.Components module

Following are Labs.Components types:


### Classes


|||
|:-----|:-----|
|[Labs.Components.ComponentAttempt](https://dev.office.com/reference/add-ins/office-mix/labs.components.componentattempt)|Base class for attempts at components.|
|[Labs.Components.ActivityComponentAttempt](https://dev.office.com/reference/add-ins/office-mix/labs.components.activitycomponentattempt)|Represents an attempt at completing an activity component.|
|[Labs.Components.ActivityComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.activitycomponentinstance)|Represents the current instance of an activity component.|
|[Labs.Components.ChoiceComponentAnswer](https://dev.office.com/reference/add-ins/office-mix/labs.components.choicecomponentanswer)|The answer to a problem presented in a choice component.|
|[Labs.Components.ChoiceComponentAttempt](https://dev.office.com/reference/add-ins/office-mix/labs.components.choicecomponentattempt)|Represents an attempt at a choice component.|
|[Labs.Components.ChoiceComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.choicecomponentinstance)|Represents an instance of a choice component.|
|[Labs.Components.ChoiceComponentResult](https://dev.office.com/reference/add-ins/office-mix/labs.components.choicecomponentresult)|The result of a choice component submission.|
|[Labs.Components.ChoiceComponentSubmission](https://dev.office.com/reference/add-ins/office-mix/labs.components.choicecomponentsubmission)|Represents the submission associated with a choice component.|
|[Labs.Components.DynamicComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.dynamiccomponentinstance)|Represents an instance of a dynamic component.|
|[Labs.Components.InputComponentAnswer](https://dev.office.com/reference/add-ins/office-mix/labs.components.inputcomponentanswer)|Represents the answer to an input component problem.|
|[Labs.Components.InputComponentAttempt](https://dev.office.com/reference/add-ins/office-mix/labs.components.inputcomponentattempt)|Represents an attempt at interacting with an input component.|
|[Labs.Components.InputComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.inputcomponentinstance)|Represents an instance of an input component.|
|[Labs.Components.InputComponentResult](https://dev.office.com/reference/add-ins/office-mix/labs.components.inputcomponentresult)|The result of an input component submission.|
|[Labs.Components.InputComponentSubmission](https://dev.office.com/reference/add-ins/office-mix/labs.components.inputcomponentsubmission)|Represents a submission to an input component.|

### Interfaces


|||
|:-----|:-----|
|[Labs.Components.IActivityComponent](https://dev.office.com/reference/add-ins/office-mix/labs.components.iactivitycomponent)|Represents an activity component. Extends [Labs.Core.IComponent](https://dev.office.com/reference/add-ins/office-mix/labs.core.icomponent).|
|[Labs.Components.IActivityComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.iactivitycomponentinstance)|Represents a specific instance of an activity component. Extends [Labs.Core.IComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.core.icomponentinstance).|
|[Labs.Components.IChoice](https://dev.office.com/reference/add-ins/office-mix/labs.components.ichoice)|An available choice for a given problem.|
|[Labs.Components.IChoiceComponent](https://dev.office.com/reference/add-ins/office-mix/labs.components.ichoicecomponent)|Enables interactions with a choice component.|
|[Labs.Components.IChoiceComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.ichoicecomponentinstance)|An instance of a choice component.|
|[Labs.Components.IDynamicComponent](https://dev.office.com/reference/add-ins/office-mix/labs.components.idynamiccomponent)|Enables interaction with a dynamic component.|
|[Labs.Components.IDynamicComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.idynamiccomponentinstance)|An instance of a dynamic component.|
|[Labs.Components.IHint](https://dev.office.com/reference/add-ins/office-mix/labs.components.ihint)|Hint for a lab problem.|
|[Labs.Components.IInputComponent](https://dev.office.com/reference/add-ins/office-mix/labs.components.iinputcomponent)|Enables interacting with an input component.|
|[Labs.Components.IInputComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.iinputcomponentinstance)|An instance of an input component.|
