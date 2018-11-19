In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.

> [!NOTE]
> This page describes an individual step of the Excel add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.

## Create the dialog page

1. Open the project in your code editor.
2. Create a file in the root of the project (where index.html is) called popup.html.
3. Add the following markup to popup.html. Note:
   - The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.
   - The markup loads a script called popup.js that you will create in a later step.
   - It also loads the Office.JS library and jQuery because they will be used in popup.js.

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css" />

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div class="padding">
                <p class="ms-font-xl">ENTER YOUR NAME</p>
            </div>
            <div class="padding">
                <input id="name-box" type="text"/>
            </div>
            <div class="padding">
                <button id="ok-button" class="ms-Button">OK</button>
            </div>
        </body>
    </html>
    ```

4. Create a file in the root of the project called popup.js.
5. Add the following code to popup.js. Note:

   - *Every page that calls APIs in the Office.JS library must first ensure that the library is fully initialized.* The best way to do that is to call the `Office.onReady()` method. If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`. For an example, see the app.js file in the project root. The call of `Office.onReady()` must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.
   - The jQuery `ready` function is called inside the `then()` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `then()` method that is chained to the call of `Office.onReady()`.

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {
                $(document).ready(function () {  

                    // TODO1: Assign handler to the OK button.

                });
            });

        // TODO2: Create the OK button handler

    }());
    ```

6. Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. Save the file.

   > [!NOTE]
   > The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane. If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the popup.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the popup.js file at all.


## Open the dialog from the task pane

1. Open the file index.html.
2. Below the `div` that contains the `freeze-header` button, add the following markup:

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. The dialog will prompt the user to enter a name and pass the user's name to the task pane. The task pane will display it in a label. Immediately below the `div` that you just added, add the following markup:

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. Open the app.js file.

5. Below the line that assigns a click handler to the `freeze-header` button, add the following code. You'll create the `openDialog` method in a later step.

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.

    ```js
    let dialog = null;
    ```

7. Below the declaration of `dialog`, add the following function. The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`. This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ```

8. Replace `TODO1` with the following code. Note:
   - The `displayDialogAsync` method opens a dialog in the center of the screen.
   - The first parameter is the URL of the page to open.
   - The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

## Process the message from the dialog and close the dialog

1. Continue in the app.js file, and replace `TODO2` with the following code. Note:
   - The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.
   - The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.
   - The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. Below the `openDialog` function, add the following function.

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## Test the add-in

1. If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.

     > [!NOTE]
     > Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.

1. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).
2. Run the command `npm start` to start a web server running on localhost.
4. Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.
6. Choose the **Open Dialog** button in the task pane.
7. While the dialog is open, drag it and resize it. Note that you can interact with the worksheet and press other buttons on the task pane. But you cannot launch a second dialog from the same task pane page.
8. In the dialog, enter a name and choose **OK**. The name appears on the task pane and the dialog closes.
9. Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.

    ![Excel tutorial - Dialog](../images/excel-tutorial-dialog-open.png)
