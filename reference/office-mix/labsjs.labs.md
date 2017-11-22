
# LabsJS.Labs

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The LabsJS.Labs module contains the set of key JavaScript APIs that you can use to create the Office Add-ins (the labs). The APIs provide the entry point for lab development.

## LabsJS.Labs API module

The Labs module contains the following types:


### Variables


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](https://dev.office.com/reference/add-ins/office-mix/labs.defaulthostbuilder)|Use this object to construct a default [Labs.Core.ILabHost](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabhost) instance.|

### Functions


|||
|:-----|:-----|
|[Labs.Connect](https://dev.office.com/reference/add-ins/office-mix/labs.connect)|Initializes a connection with the host.|
|[Labs.connect (overload)](https://dev.office.com/reference/add-ins/office-mix/labs.connect-overload)|Initializes a connection with the host and provides input parameters.|
|[Labs.isConnected](https://dev.office.com/reference/add-ins/office-mix/labs.isconnected)|Initializes a connection with the host.|
|[Labs.getConnectionInfo](https://dev.office.com/reference/add-ins/office-mix/labs.getconnectioninfo)|Retrieves configuration information associated with a specified connection.|
|[Labs.disconnect](https://dev.office.com/reference/add-ins/office-mix/labs.disconnect)|Disconnects the lab from the host and provides lab completion status.|
|[Labs.editLab](https://dev.office.com/reference/add-ins/office-mix/labs.editlab)|Opens the specified lab for editing. You can specify the lab's configuration data while in edit mode. However, you cannot edit a lab while it is being taken (that is, the lab is running).|
|[Labs.takeLab](https://dev.office.com/reference/add-ins/office-mix/labs.takelab)|Runs the specified lab and enables sending lab results to the server. Note that you cannot run a lab while it is being edited.|
|[Labs.on](https://dev.office.com/reference/add-ins/office-mix/labs.on)|Adds a new handler for a specified event..|
|[Labs.off](https://dev.office.com/reference/add-ins/office-mix/labs.off)|Removes an event handler for a specified event.|
|[Labs.getTimeline](https://dev.office.com/reference/add-ins/office-mix/labs.gettimeline)|Retrieves a [Labs.Timeline](https://dev.office.com/reference/add-ins/office-mix/labs.timeline) object instance that you can use to control the host player control.|
|[Labs.registerDeserializer](https://dev.office.com/reference/add-ins/office-mix/labs.registerdeserializer)|Deserializes a specified JSON object into an object. Should be used by component authors only.|

### Classes


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](https://dev.office.com/reference/add-ins/office-mix/labs.componentinstancebase)|Base class for the initialization of component instances.|
|[Labs.ComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.componentinstance)|Represents an instance of a component, which is an instantiation of a given component for a user at runtime. The object contains a translated view of the component for a specific run of a lab.|
|[Labs.Command](https://dev.office.com/reference/add-ins/office-mix/labs.command)|General command used to pass messages between the client and host.|
|[Labs.LabEditor](https://dev.office.com/reference/add-ins/office-mix/labs.labeditor)|The  **LabEditor** object allows you to edit a given lab as well as get and set configuration data associated with the lab.|
|[Labs.LabInstance](https://dev.office.com/reference/add-ins/office-mix/labs.labinstance)|An instance of a lab that is configured for the current user. Use this object to record and retrieve lab data for the user.|
|[Labs.Timeline](https://dev.office.com/reference/add-ins/office-mix/labs.timeline)|Provides access to the labs.js timeline feature.|
|[Labs.ValueHolder](https://dev.office.com/reference/add-ins/office-mix/labs.valueholder)|A container object that holds and tracks values for a specified lab. The value may be stored either locally or on the server.|

### Interfaces


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](https://dev.office.com/reference/add-ins/office-mix/labs.getactionscommanddata)|Allows you to retrieve data associated with a [LabsJS.Labs.Core.GetActions](https://dev.office.com/reference/add-ins/office-mix/labsjs.labs.core.getactions) command.|
|[Labs.IMessageHandler](https://dev.office.com/reference/add-ins/office-mix/labs.imessagehandler)|Interface that allows you to define event handlers.|
|[Labs.ITimelineNextMessage](https://dev.office.com/reference/add-ins/office-mix/labs.itimelinenextmessage)|Provides means for interacting with the [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx) object.|
|[Labs.SendMessageCommandData](https://dev.office.com/reference/add-ins/office-mix/labs.sendmessagecommanddata)|Data associated with a [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx) command.|
|[Labs.TakeActionCommandData](https://dev.office.com/reference/add-ins/office-mix/labs.takeactioncommanddata)|Data associated with a take action command.|

### Enumerations


|||
|:-----|:-----|
|[Labs.ConnectionState](https://dev.office.com/reference/add-ins/office-mix/labs.connectionstate)|Enumerates the possible connection states of the lab to host.|
|[Labs.ProblemState](https://dev.office.com/reference/add-ins/office-mix/labs.problemstate)|State values for a given lab.|
