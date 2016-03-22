
# Walkthrough: Creating your first lab for Office Mix
Build your first LabsJS lab using a step-by-step walkthrough.



In this walkthrough you'll create a simple LabsJS lab from scratch. Your lab will be a simple true/false quiz that provides just a single question. 

Rather than start with a Visual Studio project template, you'll start with just three empty files - this shows how straightforward a lab is: 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
You can use any code editor you want for editing these files because we do not start with a Visual Studio template. In fact, the HTML file is trivial and you if you want you can simply copy/paste the HTML markup from the tutorial files. Note, however, that it must be HTML5, so ensure that your doctype declaration is  `<!DOCTYPE html>`. The CSS file is optional. All the heavy lifting is done in the JavaScript (.js) file, TrueFalse.js.
The walkthrough is going to cover four main lab features:

- Setup (connecting to the host)
    
- Mode changes (between edit mode and view mode)
    
- Editing the lab
    
- Taking (or running) the lab
    

 **Note**  
 ---
 The file labhost.html runs on a web server and provides the hosting environment for lab development and testing. This greatly simplifies lab development. See [Get started with LabsJS for Office Mix](get-started-with-labsjs-for-office-mix.md) for information about setting up your development environment.<br/><br/>

Finally, you can view the completed JavaScript file (TrueFalse.js) among the files distributed with this SDK. What follows is a walkthrough of the coding process.

## Connecting to the lab host

Labs in this environment are able to run either with our lab host (for development and testing) or with the default runtime host provided by the Office.js host. The opening function, then, uses a simple if/else expression to test which of these hosting contexts applies.


```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

The  **PostMessageLabHost** object runs in the labhost.html development environment, whereas in production, the lab runs in PowerPoint/Office Mix using the **OfficeJSLabHost**.

Next, create a helper method to create a callback whose job is to either resolve or reject a jQuery deferred object that you pass in. Use this method,  **createCallback**, to go from the jQuery promises to the callbacks defined by labs.js.




```js
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

We also create a helper method for retrieving the lab configuration for a given question and answer.




```js
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## Mode changes

A lab is always in one of two states, or modes:  **view** and **edit**. Therefore, we need a way to capture and hold the state and behavior for the quiz; we'll create a class for this purpose.


```js
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

Additionally, we provide a helper method whose job is to update the UI of the quiz based on where the answer to a quiz question (that is, the "submission") is correct or incorrect.




```js
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

We also need a function for switching between edit and view modes.




```js
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

Our next function updates the configuration for the quiz based on change events that we've received from the UI.




```js
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

Next, we update the lab configuration data stored on the server based on the given questions and answer.




```js
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

Next we have a function that binds updates made in the lab in edit mode to the configuration changes that we've made. Following that is code for unbinding from the previously bound change handlers.




```js
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```js
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

Now comes a key part of the section, that is, methods for switching back and forth between view and edit modes. Let's start with switching from view mode to edit mode.




```js
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

And now, switching from edit mode to view mode.




```js
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

Finally, after you're connected to the host and the document is ready, start up the quiz.




```js
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## Additional resources
<a name="bk_addresources"> </a>


- [Office Mix add-ins](office-mix-add-ins.md)
    
