
# LabsJS.Labs

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The LabsJS.Labs module contains the set of key JavaScript APIs that you can use to create the Office Add-ins (the labs). The APIs provide the entry point for lab development.

## LabsJS.Labs API module

The Labs module contains the following types:


### Variables


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|Use this object to construct a default [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) instance.|

### Functions


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|Initializes a connection with the host.|
|[Labs.connect (overload)](../../reference/office-mix/labs.connect-overload.md)|Initializes a connection with the host and provides input parameters.|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|Initializes a connection with the host.|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|Retrieves configuration information associated with a specified connection.|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|Disconnects the lab from the host and provides lab completion status.|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|Opens the specified lab for editing. You can specify the lab's configuration data while in edit mode. However, you cannot edit a lab while it is being taken (that is, the lab is running).|
|[Labs.takeLab](../../reference/office-mix/labs.takelab.md)|Runs the specified lab and enables sending lab results to the server. Note that you cannot run a lab while it is being edited.|
|[Labs.on](../../reference/office-mix/labs.on.md)|Adds a new handler for a specified event..|
|[Labs.off](../../reference/office-mix/labs.off.md)|Removes an event handler for a specified event.|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|Retrieves a [Labs.Timeline](../../reference/office-mix/labs.timeline.md) object instance that you can use to control the host player control.|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|Deserializes a specified JSON object into an object. Should be used by component authors only.|

### Classes


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../reference/office-mix/labs.componentinstancebase.md)|Base class for the initialization of component instances.|
|[Labs.ComponentInstance](../../reference/office-mix/labs.componentinstance.md)|Represents an instance of a component, which is an instantiation of a given component for a user at runtime. The object contains a translated view of the component for a specific run of a lab.|
|[Labs.Command](../../reference/office-mix/labs.command.md)|General command used to pass messages between the client and host.|
|[Labs.LabEditor](../../reference/office-mix/labs.labeditor.md)|The  **LabEditor** object allows you to edit a given lab as well as get and set configuration data associated with the lab.|
|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md)|An instance of a lab that is configured for the current user. Use this object to record and retrieve lab data for the user.|
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|Provides access to the labs.js timeline feature.|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|A container object that holds and tracks values for a specified lab. The value may be stored either locally or on the server.|

### Interfaces


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|Allows you to retrieve data associated with a [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md) command.|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|Interface that allows you to define event handlers.|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|Provides means for interacting with the [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx) object.|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|Data associated with a [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx) command.|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|Data associated with a take action command.|

### Enumerations


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|Enumerates the possible connection states of the lab to host.|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)|State values for a given lab.|
