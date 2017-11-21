
# Labs.Core Interfaces
Interfaces in the  **LabsJS.Labs.Core** module

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The  **LabsJS.Labs.Core** module contains the following interfaces.

## 


|||
|:-----|:-----|
|[Labs.Core.IAction](https://dev.office.com/reference/add-ins/office-mix/labs.core.iaction)|Represents a lab action, which is an interaction that a user has with a specified lab.|
|[Labs.Core.IActionResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.iactionresult)|The results of taking an action. Depending on the type of action, the results are either be set by the server or provided by the client when the action is taken.|
|[Labs.Core.IComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.core.icomponentinstance)|Base class for instances of lab components.|
|[Labs.Core.IConfigurationInfo](https://dev.office.com/reference/add-ins/office-mix/labs.core.iconfigurationinfo)|Information about the lab configuration.|
|[Labs.Core.IConnectionResponse](https://dev.office.com/reference/add-ins/office-mix/labs.core.iconnectionresponse)|Response information returning from a connection call.|
|[Labs.Core.IGetActionOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.igetactionoptions)|Options that are passed as part of a  **get** action.|
|[Labs.Core.ILabCreationOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabcreationoptions)|Options that are passed as part of a lab create operation.|
|[Labs.Core.ILabHostVersionInfo](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabhostversioninfo)|Version information about the lab host.|
|[Labs.Core.IActionOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.iactionoptions)|Definition of lab action options. The options that are passed when performing a given action.|
|[Labs.Core.IUserInfo](https://dev.office.com/reference/add-ins/office-mix/labs.core.iuserinfo)|Provides user information relevant to the lab.|
|[Labs.Core.IValueInstance](https://dev.office.com/reference/add-ins/office-mix/labs.core.ivalueinstance)|An [Labs.Core.IValue](https://dev.office.com/reference/add-ins/office-mix/labs.core.ivalue) object instance that contains value data, if any.|
|[Labs.Core.IVersion](https://dev.office.com/reference/add-ins/office-mix/labs.core.iversion)|Provides the lab version information.|
|[Labs.Core.IAnalyticsConfiguration](https://dev.office.com/reference/add-ins/office-mix/labs.core.ianalyticsconfiguration)|Custom analytics configuration information. Allows you to specify which IFrame to load to display custom analytics for a user's run of a lab.|
|[Labs.Core.ICompletionStatus](https://dev.office.com/reference/add-ins/office-mix/labs.core.icompletionstatus)|Completion status for the lab. The status is passed when completing the lab to indicate the result of the interaction.|
|[Labs.Core.ILabCallback](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabcallback)|The interface for handling Labs.js callback methods.|
|[Labs.Core.ILabObject](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabobject)|An object associated with a lab. The object contains a type field that indicates what type of object it is.|
|[Labs.Core.ITimelineConfiguration](https://dev.office.com/reference/add-ins/office-mix/labs.core.itimelineconfiguration)|Configuration options for the [Labs.Timeline](https://dev.office.com/reference/add-ins/office-mix/labs.timeline). Allows you to specify a set of timeline configuration options.|
|[Labs.Core.IUserData](https://dev.office.com/reference/add-ins/office-mix/labs.core.iuserdata)|The base interface to represent custom user data that is stored on an object.|
|[Labs.Core.IValue](https://dev.office.com/reference/add-ins/office-mix/labs.core.ivalue)|Base class for values stored with a lab.|
|[Labs.Core.IConfiguration](https://dev.office.com/reference/add-ins/office-mix/labs.core.iconfiguration)|Lab configuration data structure.|
|[Labs.Core.IConfigurationInstance](https://dev.office.com/reference/add-ins/office-mix/labs.core.iconfigurationinstance)|Base class for instances of a lab configuration.|
|[Labs.Core.IComponent](https://dev.office.com/reference/add-ins/office-mix/labs.core.icomponent)|Base class for representing components of a lab.|
|[Labs.Core.ILabHost](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabhost)|Provides an abstraction layer for connecting Labs.js to the host.|
|[Labs.Core.ModeChangedEventData](https://dev.office.com/reference/add-ins/office-mix/labs.core.modechangedeventdata)|Data associated with a mode changed event.|
|[Labs.Core.IEventCallback](https://dev.office.com/reference/add-ins/office-mix/labs.core.ieventcallback)|Interface for handling EventManager callbacks.|
