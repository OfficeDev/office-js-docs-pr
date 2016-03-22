
# Guidelines for creating labs for Mix using LabsJS



The LabsJS library (labs.js) supports writing specialized Office Add-ins (called labs) that integrate with Office Mix. Office Mix then renders the labs using Microsoft PowerPoint. While we call these components "labs," let's be clear that what we're creating are special Office Add-ins that are Office Mix Add-ins.

The LabsJS content helps you implementing the labs.js JavaScript API by providing guidance and examples. This library is built on top of the [JavaScript API for Office](../../../reference/javascript-api-for-office.md) (Office.js) and provides an abstraction layer that is optimized for add-ins embedded in Office Mix.


## General guidelines


The following are some general guidelines to help when writing add-ins using the LabJS API.


### Scripts

Because the labs.js library is an abstraction layer on office.js, and therefore has a dependency on office.js, both the office.js and labs.js library files must be included in your development projects. 

You can reference the office.js library here:  `<script src="https://sforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>`.

The labs.js library is included with the LabsJS SDK. Alternatively, you can reference the labs.js library on a CDN at  <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Note that the production version of your lab must reference the version stored on the CDN.


 >**Note**:  In addition to the JavaScript file (labs-1.0.4.js), we provide a TypeScript definition file of the labs API (labs-1.0.4.d.ts). The definition file was built against TypeScript version 0.9.1.1.


### Callbacks and error handling

Several methods in the labs.js API operate asynchronously. For these operations, the API adopts a standard callback interface,  **ILabCallback**. 


```js
function(err, result) {
}
```

The callback method takes two parameters,  _err_ and _result_. The  _err_ field remains **null** unless there is an error. The _result_ field returns the result of the operation.

The callback operation never fires immediately, even if the result is available immediately. Instead, it fires on a separate execution of the JavaScript event loop (by way of the  **setTimeout** call). By adopting this callback definition, you can easily integrate labs.js with your promise API of choice. For example, you can substitute jQuery promises for these callbacks with a simple translation method, as shown in the following example.




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### Lab host and DefaultLabHost

The lab host ( **ILabHost**) is the underlying driver that supports the development of Labs. By default, this is set to a host that integrates with office.js.

For testing purposes, and to run your lab within labhost.html, you need to switch out to a host that works in the simulation environment. The following code example shows how to do this using a query parameter. Alternatively, you can change  **DefaultHostBuilder** to integrate your Lab add-in with a different platform altogether.




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### Initialization

Initialization establishes the communication pathway between the lab and its host. Initialize your lab by calling the following.


```js
Labs.connect((err, connectionResponse) => {});
```

After you initialize, you can call other methods of the labs.js API. The  _connectionResponse_ parameter contains information about the host, user, and other connection-related information. For more information about the values returned, see the [Labs.Core.IConnectionResponse](../../../reference/office-mix/labs.core.iconnectionresponse.md).


### Time format

Labs.js stores numbers as milliseconds elapsed since January 1st 1970 UTC. This matches date format of the JavaScript [Date object](http://msdn.microsoft.com/en-us/library/ie/cd9w2te4%28v=vs.94%29.aspx),


### Timeline

The lab can also interact with the lesson player timeline. The timeline allows the lab to tell the lesson player to advance to the next slide. The timeline object is retrieved by calling the  **Labs.getTimeline** method.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## Handling events


The LabsJS events API tracks lab-specific events and enables you to add event handlers so you can respond to or act on the events. The event methods, of which there are three, are on the  **EventTypes** object: **ModeChanged**,  **Activate**, and  **Deactivate**. 


### Mode change

The  **ModeChanged** event fires when the specified lab changes from edit mode to view mode. Edit mode is visible when the lab is viewed in PowerPoint edit mode. View mode is visible when PowerPoint is rendering the slide show or when the lab is being displayed in the Office Mix lesson player. View mode should always display what the user sees when taking the lab. Edit mode allows the user to configure the lab.

Data in the  **ModeChangedEventData** object that is passed to the callback contains information about the current mode. The following code shows how to use the **ModeChanged** event.




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### Activate

The  **activate** event fires when the PowerPoint slide that the lab is on becomes active in the lesson player.


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### Deactivate

The  **deactivate** event fires when the PowerPoint slide the lab is on is no longer the active slide.


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### Timeline

The lab can also interact with the lesson player timeline. The timeline allows the lab to tell the lesson player to advance to the next slide. The timeline object is retrieved by calling the  **Labs.getTimeline** method.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## Additional resources



- [Office Mix add-ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
