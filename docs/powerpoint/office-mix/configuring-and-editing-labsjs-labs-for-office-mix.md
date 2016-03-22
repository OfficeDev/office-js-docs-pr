
# Configuring and editing LabsJS labs for Office Mix



Office Mix provides office.js methods to get and set lab configurations. The configuration indicates to Office Mix what type of lab you are creating, as well as what type of data the lab will send back. This information is used to collect and visualize analytics.

## Getting the lab editor

The lab editor, the [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) object, allows you to edit your lab and get and set your lab configuration. When you have finished editing your lab, you must call the **Done** method. However, calling the **Done** method is not required except when you are trying to take or run a lab that you are editing. Note that only one instance of the lab can be opened at a time.

The following code shows how to get the lab editor.




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

Use the  **getConfiguration** and **setConfiguration** methods on [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) to store the configuration for a given lab. The configuration ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) indicates to Office Mix what data will be collected and processed by the lab. A configuration contains general information about a lab, including the name, version, and other configuration options. The most important part of the configuration is the definition of the lab components.

The following code shows how to set and get a configuration. To set a configuration, simply create the configuration object, and then call the  **setConfiguration** method. To then retrieve the configuration, you call the **getConfiguration** method on the lab editor object.




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## Closing the editor

To close the editor, call the  **Done** method on the editor when you're finished editing the lab. Note that you cannot both take and edit a lab. After you have called **Done**, however, you can then either edit or run the lab.


## Interacting with a lab

After you have set the lab configuration, you are ready to begin interacting with the lab. When the lab runs inside PowerPoint, interactions are simulated. When the lab runs inside the Office Mix lesson player, however, the data is stored in the Office Mix database and used in analytics.


### Getting the lab instance

You interact with the lab using the [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md) object, which is an instance of the configured lab for the current user. To run (or "take") the lab, call the [Labs.takeLab](../../../reference/office-mix/labs.takelab.md) function.


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

The instance object contains an array of component instances ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md), [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)) that map to the components that you specified in the configuration. In fact, an instance is simply a transformed version of the configuration that is used to attach server side IDs to instance objects, as well as to hide certain fields from the user when applicable (for example, hints, answers, and so on).


### Managing state

State is temporary storage associated with a user running a given lab. You can use the store to persist information between successive invocations of the lab. For example, a programming lab could store the user's current work in progress.

To  **set** state, use the following code.




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

To  **get** state, use the following code.




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## Component instances and results

What follows is an overview of how to implement the instances of the four component types, as well as brief examples of the component methods. 

First, however, you need to familiarize yourself with two core concepts when working with components instances. The first of these is the concept of  **attempts** and **values**.

 **Attempts**

An attempt is a try by a user to complete a component instance. For example, in the case of a multiple choice question, an attempt starts when the user begins to work the problem and it ends when a final score is assigned. The Office Mix analytics then aggregate user results for the problem.


 >**Note**:  Attempts can be used for all component types except for the  **DynamicComponent** type.

You can retrieve the results for all the attempts associated with a given component instance by using the  **getAttempts** method. After retrieving the results, the user can either re-try one of the existing attempts by using the **resume** method, or create a new attempt by using the **createAttempt** method. The following example shows the process.




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **Values**

Component instances contain a dictionary of keys that map to an array of values. You can use the array to store hints, feedback, or any other set of values that you want to associate with the component. The component instance provides access to these values using the  **getValues** method.

Querying for a hint value, for example, causes the analytics to mark that the user took a hint. Values are tracked on a per-attempt basis.

The following code example shows how to query for a hint.




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### ActivityComponentInstance


Use the  **ActivityComponentInstace** object to track a user's interaction with an activity component. This class provides a **complete** method to indicate that the user has finished interacting with the activity. The method can indicate that the user has completed an assigned task, finished a reading, or any other end point associated with the activity. The following code shows how to use the **complete** method.


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### ChoiceComponentInstance


Use the  **ChoiceComponentInstance** object to track a user's interaction with a choice component. Choice components are problems that present the user with a list of choices that they then need to select from. There may or may not be a correct answer. The class provides two primary methods: **getSubmissions** and **submit**. The  **getSubmissions** method allows you to retrieve previously stored submissions; the **submit** method allows a new submission to be stored. The following code examples illustrate using the methods.


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### InputComponentInstance


Use the  **InputComponentInstance** object to track a user's interaction with an input component. The class provides two primary methods: **getSubmission** and **submit**. The  **getSubmissions** method allows you to retrieve previously stored submissions; the **submit** method allows you to store a new submission. The following code snippet illustrates using the **getSubmissions** method.


```js
var submissions = this._attempt.getSubmissions();
```

When using the  **submit** method, note that the **InputComponentAnswer** object represents the submitted answer, and the **InputComponentResult** object contains the result. The return value is a **InputComponentSubmission** object that contains the answer, result, and a timestamp that indicates when the result was submitted.




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### DynamicComponentInstance


Use the  **DynamicComponentInstance** object to track a user's interaction with a dynamic component. The primary methods in this class are **getComponents**,  **createComponent**, and  **close**.

The  **getComponents** method allows you to retrieve a list of previously created component instance, as shown in the following example.




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

The  **createComponent** method constructs a new component and returns that component instance, as shown in the following example.




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

Use the  **close** method to indicate that you have finished using the dynamic component to create new components. Note that you can also use an **isClosed** Boolean method to test whether the dynamic component instance has been closed. The following code shows how to use the **close** method.




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## Additional resources



- [Office Mix add-ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Walkthrough: Creating your first lab for Office Mix](../../powerpoint/office-mix/walkthrough:-creating-your-first-lab-for-office-mix.md)
    
