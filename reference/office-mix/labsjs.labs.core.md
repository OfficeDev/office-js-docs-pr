
# LabsJS.Labs.Core
Provides a high-level view of the LabsJS.Labs.Core JavaScript API reference.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The core interfaces, data structures, and classes are components that are shared by LabsJS and the presentation driver (in this case, Office Mix), to create an interactive bridge between the two.

## LabsJS.Labs.Core API module

The Labs.Core module contains the following types:


### Classes


|||
|:-----|:-----|
|[Labs.Core.Permissions](../../reference/office-mix/labs.core.permissions.md)|Static class representing the permissions enabled for a given user of the lab.|

### Interfaces


|||
|:-----|:-----|
|[Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)|Represents a lab action, which is an interaction that a user has with a specified lab.|
|[Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md)|The results of taking an action. Depending on the type of action, the results are either be set by the server or provided by the client when the action is taken.|
|[Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md)|Base class for instances of lab components.|
|[Labs.Core.IConfigurationInfo](../../reference/office-mix/labs.core.iconfigurationinfo.md)|Information about the lab configuration.|
|[Labs.Core.IConnectionResponse](../../reference/office-mix/labs.core.iconnectionresponse.md)|Response information returning from a connection call.|
|[Labs.Core.IGetActionOptions](../../reference/office-mix/labs.core.igetactionoptions.md)|Options that are passed as part of a  **get** action.|
|[Labs.Core.ILabCreationOptions](../../reference/office-mix/labs.core.ilabcreationoptions.md)|Options that are passed as part of a lab create operation.|
|[Labs.Core.ILabHostVersionInfo](../../reference/office-mix/labs.core.ilabhostversioninfo.md)|Version information about the lab host.|
|[Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md)|Definition of lab action options. The options that are passed when performing a given action.|
|[Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)|Provides user information relevant to the lab.|
|[Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)|An [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) object instance that contains value data, if any.|
|[Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)|Provides the lab version information.|
|[Labs.Core.IAnalyticsConfiguration](../../reference/office-mix/labs.core.ianalyticsconfiguration.md)|Custom analytics configuration information. Allows you to specify which IFrame to load to display custom analytics for a user's run of a lab.|
|[Labs.Core.ICompletionStatus](../../reference/office-mix/labs.core.icompletionstatus.md)|Completion status for the lab. The status is passed when completing the lab to indicate the result of the interaction.|
|[Labs.Core.ILabCallback](../../reference/office-mix/labs.core.ilabcallback.md)|The interface for handling Labs.js callback methods.|
|[Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)|An object associated with a lab. The object contains a type field that indicates what type of object it is.|
|[Labs.Core.ITimelineConfiguration](../../reference/office-mix/labs.core.itimelineconfiguration.md)|Configuration options for the [Labs.Timeline](../../reference/office-mix/labs.timeline.md). Allows you to specify a set of timeline configuration options.|
|[Labs.Core.IUserData](../../reference/office-mix/labs.core.iuserdata.md)|The base interface to represent custom user data that is stored on an object.|
|[Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md)|Base class for values stored with a lab.|
|[Labs.Core.IConfiguration](../../reference/office-mix/labs.core.iconfiguration.md)|Lab configuration data structure.|
|[Labs.Core.IConfigurationInstance](../../reference/office-mix/labs.core.iconfigurationinstance.md)|Base class for instances of a lab configuration.|
|[Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)|Base class for representing components of a lab.|
|[Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)|Provides an abstraction layer for connecting Labs.js to the host.|
|[Labs.Core.ModeChangedEventData](../../reference/office-mix/labs.core.modechangedeventdata.md)|Data associated with a mode changed event.|
|[Labs.Core.IEventCallback](../../reference/office-mix/labs.core.ieventcallback.md)|Interface for handling EventManager callbacks.|

### Enumerations


|||
|:-----|:-----|
|[Labs.Core.LabMode](../../reference/office-mix/labs.core.labmode.md)|Values denoting the current state of the lab.|
