---
title: 'Tutorial: Build a message compose Outlook add-in'
description: In this tutorial, you will build an Outlook add-in that inserts GitHub gists into the body of a new message.
ms.date: 12/09/2025
ms.service: outlook
#Customer intent: As a developer, I want to create a message compose Outlook add-in.
ms.localizationpriority: high
---

# Tutorial: Build a message compose Outlook add-in

This tutorial teaches you how to build an Outlook add-in that can be used in message compose mode to insert content into the body of a message.

In this tutorial, you will:

> [!div class="checklist"]
>
> - Create an Outlook add-in project
> - Define buttons that appear in the compose message window
> - Implement a first-run experience that collects information from the user and fetches data from an external service
> - Implement a UI-less button that invokes a function
> - Implement a task pane that inserts content into the body of a message

> [!TIP]
> If you want a completed version of this tutorial, visit the [Office Add-ins samples repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/outlook-tutorial).

## Prerequisites

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Visual Studio Code (VS Code)](https://code.visualstudio.com/) or your preferred code editor.

- Outlook on the web, [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), or Outlook 2016 or later on Windows (connected to a Microsoft 365 account).

- A [GitHub](https://www.github.com) account.

## Setup

The add-in that you'll create in this tutorial will read [gists](https://gist.github.com) from the user's GitHub account and add the selected gist to the body of a message. Complete the following steps to create two new gists that you can use to test the add-in you're going to build.

1. [Login to GitHub](https://github.com/login).

1. [Create a new gist](https://gist.github.com).

    - In the **Gist description...** field, enter **Hello World Markdown**.

    - In the **Filename including extension...** field, enter **test.md**.

    - Add the following markdown to the multiline textbox.

        ```markdown
        # Hello World

        This is content converted from Markdown!

        Here's a JSON sample:

          ```json
          {
            "foo": "bar"
          }
          ```
        ```

    - Select the **Create public gist** button.

1. [Create another new gist](https://gist.github.com).

    - In the **Gist description...** field, enter **Hello World Html**.

    - In the **Filename including extension...** field, enter **test.html**.

    - Add the following markdown to the multiline textbox.

        ```HTML
        <html>
          <head>
            <style>
            h1 {
              font-family: Calibri;
            }
            </style>
          </head>
          <body>
            <h1>Hello World!</h1>
            <p>This is a test</p>
          </body>
        </html>
        ```

    - Select the **Create public gist** button.

## Create an Outlook add-in project

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

1. The steps to create the project vary slightly depending on the type of manifest.

    [!INCLUDE [Unified manifest value proposition and feedback request](../includes/unified-manifest-value-prop-feedback.md)]

    # [Unified manifest for Microsoft 365](#tab/jsonmanifest)

    - **Choose a project type** - `Office Add-in Task Pane project`

    - **Choose a script type** - `JavaScript`

    - **What do you want to name your add-in?** - `Git the gist`

    - **Which Office client application would you like to support?** - `Outlook`

    - **Which manifest would you like to use?** - `unified manifest for Microsoft 365`

    :::image type="content" source="../images/yo-office-outlook-json-manifest-javascript-gist.png" alt-text="The prompts and answers for the Yeoman generator with unified manifest and JavaScript options chosen.":::

    # [Add-in only manifest](#tab/xmlmanifest)

    - **Choose a project type** - `Office Add-in Task Pane project`

    - **Choose a script type** - `JavaScript`

    - **What do you want to name your add-in?** - `Git the gist`

    - **Which Office client application would you like to support?** - `Outlook`

    - **Which manifest would you like to use?** - `Add-in only manifest`

    :::image type="content" source="../images/yo-office-outlook-xml-manifest-javascript-gist.png" alt-text="The prompts and answers for the Yeoman generator in a command line interface.":::

    ---

    After you complete the wizard, the generator creates the project and installs supporting Node components.

1. Navigate to the root directory of the project.

    ```command&nbsp;line
    cd "Git the gist"
    ```

1. This add-in uses the following libraries.

    - [Showdown](https://github.com/showdownjs/showdown) library to convert Markdown to HTML.
    - [URI.js](https://github.com/medialize/URI.js) library to build relative URLs.
    - [jQuery](https://jquery.com/) library to simplify DOM interactions.

     To install these tools for your project, run the following command in the root directory of the project.

    ```command&nbsp;line
    npm install showdown urijs jquery --save
    ```

1. Open your project in VS Code or your preferred code editor.

    [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

### Update the manifest

The manifest for an add-in controls how it appears in Outlook. It defines the way the add-in appears in the add-in list and the buttons that appear on the ribbon, and it sets the URLs for the HTML and JavaScript files used by the add-in.

#### Specify basic information

Make the following updates in the manifest file to specify some basic information about the add-in.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

1. Locate the `"description"` property, replace the default `"short"` and `"full"` values with descriptions of the add-in, and save the file.

    ```json
    "description": {
        "short": "Gets gists.",
        "full": "Allows users to access their GitHub gists."
    },
    ```

1. Save the file.

# [Add-in only manifest](#tab/xmlmanifest)

1. Locate the `<ProviderName>` element and replace the default value with your company name.

    ```xml
    <ProviderName>Contoso</ProviderName>
    ```

1. Locate the `<Description>` element, replace the default value with a description of the add-in, and save the file.

    ```xml
    <Description DefaultValue="Allows users to access their GitHub gists."/>
    ```

---

#### Test the generated add-in

Before going any further, let's test the basic add-in that the generator created to confirm that the project is set up correctly.

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Run the following command in the root directory of your project. When you run this command, the local web server starts and your add-in is sideloaded.

    ```command&nbsp;line
    npm start
    ```

    [!INCLUDE [outlook-manual-sideloading](../includes/outlook-manual-sideloading.md)]

1. In Outlook, open an existing message and select the **Show Taskpane** button.

1. When prompted with the **WebView Stop On Load** dialog box, select **OK**.

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

    If everything's been set up correctly, the task pane opens and renders the add-in's welcome page.

    :::image type="content" source="../images/button-and-pane.png" alt-text="The Show Taskpane button and Git the gist task pane added by the sample.":::

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-outlook-dev-add-in.md)]

## Define buttons

Now that you've verified the base add-in works, you can customize it to add more functionality. By default, the manifest only defines buttons for the read message window. Let's update the manifest to remove the buttons from the read message window and define two new buttons for the compose message window:

- **Display gist list**: a button that opens a task pane

- **Insert default gist**: a button that invokes a function

The procedure depends on which manifest you're using.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

Take the following steps:

1. Open the **manifest.json** file.

1. In the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) array, there are two runtime objects. For the second one, with the `"id"` of `"CommandsRuntime"`, change the [`"actions.id"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#id) to `"insertDefaultGist"`. This is the name of a function that you create in a later step. When you're done, the runtime object should look like the following:

    ```json
    {
        "id": "CommandsRuntime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "actions": [
            {
                "id": "insertDefaultGist",
                "type": "executeFunction"
            }
        ]
    }
    ```

1. Change the item in the [`"extensions.ribbons.contexts"`](/microsoft-365/extensibility/schema/extension-ribbons-array#contexts) array to `"mailCompose"`. This means the buttons will appear only in a new message or reply window.

    ```json
    "contexts": [
        "mailCompose"
    ],
    ```

1. The [`"extensions.ribbons.tabs.groups"`](/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item#groups) array has a group object in it. Make the following changes to this object.

    1. Change the `"id"` property to `"msgComposeCmdGroup"`.
    1. Change the `"label"` property to "Git the gist".

1. That same group object has a `"controls"` array with two control objects. We need to make changes to the JSON for each of them. In the first one, take these steps.

    1. Change the `"id"` to `"msgComposeShowGistListTaskPane"`.
    1. Change the `"label"` to "Display gist list".
    1. Change the `"supertip.title"` to "Display gist list".
    1. Change the `"supertip.description"` to "Displays a list of your gists and allows you to insert their contents into the current message."

1. In the second control object, take these steps.

    1. Change the `"id"` to `"msgComposeInsertDefaultGist"`.
    1. Change the `"label"` to "Insert default gist".
    1. Change the `"supertip.title"` to "Insert default gist".
    1. Change the `"supertip.description"` to "Inserts the content of the gist you mark as default into the current message."
    1. Change the `"actionId"` to `"insertDefaultGist"`. This matches the `"action.id"` of the `"CommandsRuntime"` that you set in an earlier step.

    When you're done, the [`"ribbons"`](/microsoft-365/extensibility/schema/element-extensions#ribbons) property should look like the following:

    ```json
    "ribbons": [
        {
            "contexts": [
                "mailCompose"
            ],
            "tabs": [
                {
                    "builtInTabId": "TabDefault",
                    "groups": [
                        {
                            "id": "msgComposeCmdGroup",
                            "label": "Git the gist",
                            "icons": [
                                {
                                    "size": 16,
                                    "file": "https://localhost:3000/assets/icon-16.png"
                                },
                                {
                                    "size": 32,
                                    "file": "https://localhost:3000/assets/icon-32.png"
                                },
                                {
                                    "size": 80,
                                    "file": "https://localhost:3000/assets/icon-80.png"
                                }
                            ],
                            "controls": [
                                {
                                    "id": "msgComposeInsertGist",
                                    "type": "button",
                                    "label": "Display gist list",
                                    "icons": [
                                        {
                                            "size": 16,
                                            "file": "https://localhost:3000/assets/icon-16.png"
                                        },
                                        {
                                            "size": 32,
                                            "file": "https://localhost:3000/assets/icon-32.png"
                                        },
                                        {
                                            "size": 80,
                                            "file": "https://localhost:3000/assets/icon-80.png"
                                        }
                                    ],
                                    "supertip": {
                                        "title": "Display gist list",
                                        "description": "Displays a list of your gists and allows you to insert their contents into the current message."
                                    },
                                    "actionId": "TaskPaneRuntimeShow"
                                },
                                {
                                    "id": "msgComposeInsertDefaultGist",
                                    "type": "button",
                                    "label": "Insert default gist",
                                    "icons": [
                                        {
                                            "size": 16,
                                            "file": "https://localhost:3000/assets/icon-16.png"
                                        },
                                        {
                                            "size": 32,
                                            "file": "https://localhost:3000/assets/icon-32.png"
                                        },
                                        {
                                            "size": 80,
                                            "file": "https://localhost:3000/assets/icon-80.png"
                                        }
                                    ],
                                    "supertip": {
                                        "title": "Insert default gist",
                                        "description": "Inserts the content of the gist you mark as default into the current message."
                                    },
                                    "actionId": "insertDefaultGist"
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
    ```

1. Save your changes to the manifest.

# [Add-in only manifest](#tab/xmlmanifest)

### Remove the MessageReadCommandSurface extension point

1. Open the **manifest.xml** file.

1. Locate the `<ExtensionPoint>` element with type **MessageReadCommandSurface** and delete it (including its closing tag). This removes the add-in buttons from the read message window.

### Add the MessageComposeCommandSurface extension point

1. In **manifest.xml**, locate the line in the manifest that reads `</DesktopFormFactor>`.

1. Immediately before this line, insert the following XML markup. Note the following about this markup.

    - The `<ExtensionPoint>` with `xsi:type="MessageComposeCommandSurface"` indicates that you're defining buttons to add to the compose message window.

    - By using an `<OfficeTab>` element with `id="TabDefault"`, you're indicating you want to add the buttons to the default tab on the ribbon.

    - The `<Group>` element defines the grouping for the new buttons, with a label set by the **groupLabel** resource.

    - The first `<Control>` element contains an `<Action>` element with `xsi:type="ShowTaskPane"`, so this button opens a task pane.

    - The second `<Control>` element contains an `<Action>` element with `xsi:type="ExecuteFunction"`, so this button invokes a JavaScript function contained in the function file.

    ```xml
    <!-- Message Compose -->
    <ExtensionPoint xsi:type="MessageComposeCommandSurface">
      <OfficeTab id="TabDefault">
        <Group id="msgComposeCmdGroup">
          <Label resid="GroupLabel"/>
          <Control xsi:type="Button" id="msgComposeShowGistListTaskPane">
            <Label resid="TaskpaneButton.Label"/>
            <Supertip>
              <Title resid="TaskpaneButton.Title"/>
              <Description resid="TaskpaneButton.Tooltip"/>
            </Supertip>
            <Icon>
              <bt:Image size="16" resid="Icon.16x16"/>
              <bt:Image size="32" resid="Icon.32x32"/>
              <bt:Image size="80" resid="Icon.80x80"/>
            </Icon>
            <Action xsi:type="ShowTaskpane">
              <SourceLocation resid="Taskpane.Url"/>
            </Action>
          </Control>
          <Control xsi:type="Button" id="msgComposeInsertDefaultGist">
            <Label resid="FunctionButton.Label"/>
            <Supertip>
              <Title resid="FunctionButton.Title"/>
              <Description resid="FunctionButton.Tooltip"/>
            </Supertip>
            <Icon>
              <bt:Image size="16" resid="Icon.16x16"/>
              <bt:Image size="32" resid="Icon.32x32"/>
              <bt:Image size="80" resid="Icon.80x80"/>
            </Icon>
            <Action xsi:type="ExecuteFunction">
              <FunctionName>insertDefaultGist</FunctionName>
            </Action>
          </Control>
        </Group>
      </OfficeTab>
    </ExtensionPoint>
    ```

### Update resources in the manifest

The previous code references labels, tooltips, and URLs that you need to define before the manifest will be valid. You'll specify this information in the `<Resources>` section of the manifest.

1. In **manifest.xml**, locate the `<Resources>` element in the manifest file and delete the entire element (including its closing tag).

1. In that same location, add the following markup to replace the `<Resources>` element you just removed.

    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Git the gist"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Display gist list"/>
        <bt:String id="TaskpaneButton.Title" DefaultValue="Display gist list"/>
        <bt:String id="FunctionButton.Label" DefaultValue="Insert default gist"/>
        <bt:String id="FunctionButton.Title" DefaultValue="Insert default gist"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Displays a list of your gists and allows you to insert their contents into the current message."/>
        <bt:String id="FunctionButton.Tooltip" DefaultValue="Inserts the content of the gist you mark as default into the current message."/>
      </bt:LongStrings>
    </Resources>
    ```

1. Save your changes to the manifest.

---

### Reinstall the add-in

You must reinstall the add-in for the manifest changes to take effect.

1. If the web server is running, run the following command.

    ```command&nbsp;line
    npm stop
    ```

1. Run the following command to start the local web server and automatically sideload your add-in.

    ```command&nbsp;line
    npm start
    ```

After you've reinstalled the add-in, you can verify that it installed successfully by checking for the commands **Display gist list** and **Insert default gist** in a compose message window. Note that nothing will happen if you select either of these items, because you haven't yet finished building this add-in.

- If you're running this add-in in Outlook 2016 or later on Windows, you should see two new buttons on the ribbon of the compose message window: **Display gist list** and **Insert default gist**.

  :::image type="content" source="../images/add-in-buttons-in-windows.png" alt-text="The add-in buttons as they appear in classic Outlook on Windows.":::

- If you're running this add-in in Outlook on the web or new Outlook on Windows, select **Apps** from the ribbon of the compose message window, then select **Git the gist** to see the **Display gist list** and **Insert default gist** options.

  :::image type="content" source="../images/add-in-buttons-in-owa.png" alt-text="The message compose form in Outlook on the web with the add-in button and pop-up menu highlighted.":::

## Implement a first-run experience

This add-in needs to be able to read gists from the user's GitHub account and identify which one the user has chosen as the default gist. In order to achieve these goals, the add-in must prompt the user to provide their GitHub username and choose a default gist from their collection of existing gists. Complete the steps in this section to implement a first-run experience that displays a dialog to collect this information from the user.

### Create the UI of the dialog

Let's start by creating the UI for the dialog.

1. Within the **./src** folder, create a new subfolder named **settings**.

1. In the **./src/settings** folder, create a file named **dialog.html**.

1. In **dialog.html**, add the following markup to define a basic form with a text input for a GitHub username and an empty list for gists that'll be populated via JavaScript.

    ```html
    <!DOCTYPE html>
    <html>
    
    <head>
      <meta charset="UTF-8" />
      <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
      <title>Settings</title>
    
      <!-- Office JavaScript API -->
      <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    
    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
      <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css"/>
    
      <!-- Template styles -->
      <link href="dialog.css" rel="stylesheet" type="text/css" />
    </head>
    
    <body class="ms-font-l">
      <main>
        <section class="ms-font-m ms-fontColor-neutralPrimary">
          <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning">
            <div class="ms-MessageBar-content">
              <div class="ms-MessageBar-icon">
                <i class="ms-Icon ms-Icon--Info"></i>
              </div>
              <div class="ms-MessageBar-text">
                Oops! It looks like you haven't configured <strong>Git the gist</strong> yet.
                <br/>
                Please configure your GitHub username and select a default gist, then try that action again!
              </div>
            </div>
          </div>
          <div class="ms-font-xxl">Settings</div>
          <div class="ms-Grid">
            <div class="ms-Grid-row">
              <div class="ms-TextField">
                <label class="ms-Label">GitHub Username</label>
                <input class="ms-TextField-field" id="github-user" type="text" value="" placeholder="Please enter your GitHub username">
              </div>
            </div>
            <div class="error-display ms-Grid-row">
              <div class="ms-font-l ms-fontWeight-semibold">An error occurred:</div>
              <pre><code id="error-text"></code></pre>
            </div>
            <div class="gist-list-container ms-Grid-row">
              <div class="list-title ms-font-xl ms-fontWeight-regular">Choose Default Gist</div>
              <form>
                <div id="gist-list">
                </div>
              </form>
            </div>
          </div>
          <div class="ms-Dialog-actions">
            <div class="ms-Dialog-actionsRight">
              <button class="ms-Dialog-action ms-Button ms-Button--primary" id="settings-done" disabled>
                <span class="ms-Button-label">Done</span>
              </button>
            </div>
          </div>
        </section>
      </main>
      <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
      <script type="text/javascript" src="../helpers/gist-api.js"></script>
      <script type="text/javascript" src="dialog.js"></script>
    </body>
    
    </html>
    ```

    You may have noticed that the HTML file references a JavaScript file, **gist-api.js**, that doesn't yet exist. This file will be created in the [Fetch data from GitHub](#fetch-data-from-github) section below.

1. Save your changes.

1. Next, create a file in the **./src/settings** folder named **dialog.css**.

1. In **dialog.css**, add the following code to specify the styles that are used by **dialog.html**.

    ```css
    section {
      margin: 10px 20px;
    }
    
    .not-configured-warning {
      display: none;
    }
    
    .error-display {
      display: none;
    }
    
    .gist-list-container {
      margin: 10px -8px;
      display: none;
    }
    
    .list-title {
      border-bottom: 1px solid #a6a6a6;
      padding-bottom: 5px;
    }
    
    ul {
      margin-top: 10px;
    }
    
    .ms-ListItem-secondaryText,
    .ms-ListItem-tertiaryText {
      padding-left: 15px;
    }
    ```

1. Save your changes.

### Develop the functionality of the dialog

Now that you've defined the dialog UI, you can write the code that makes it actually do something.

1. In the **./src/settings** folder, create a file named **dialog.js**.

1. Add the following code. Note that this code uses jQuery to register events and uses the `messageParent` method to send the user's choices back to the caller.

    ```js
    (function() {
      'use strict';
    
      // The onReady function must be run each time a new page is loaded.
      Office.onReady(function() {
        $(document).ready(function() {
          if (window.location.search) {
            // Check if warning should be displayed.
            const warn = getParameterByName('warn');
    
            if (warn) {
              $('.not-configured-warning').show();
            } else {
              // See if the config values were passed.
              // If so, pre-populate the values.
              const user = getParameterByName('gitHubUserName');
              const gistId = getParameterByName('defaultGistId');
    
              $('#github-user').val(user);
              loadGists(user, function(success) {
                if (success) {
                  $('.ms-ListItem').removeClass('is-selected');
                  $('input').filter(function() {
                    return this.value === gistId;
                  }).addClass('is-selected').attr('checked', 'checked');
                  $('#settings-done').removeAttr('disabled');
                }
              });
            }
          }
    
          // When the GitHub username changes,
          // try to load gists.
          $('#github-user').on('change', function() {
            $('#gist-list').empty();
            const ghUser = $('#github-user').val();
    
            if (ghUser.length > 0) {
              loadGists(ghUser);
            }
          });
    
          // When the Done button is selected, send the
          // values back to the caller as a serialized
          // object.
          $('#settings-done').on('click', function() {
            const settings = {};
            settings.gitHubUserName = $('#github-user').val();
            const selectedGist = $('.ms-ListItem.is-selected');
    
            if (selectedGist) {
              settings.defaultGistId = selectedGist.val();
              sendMessage(JSON.stringify(settings));
            }
          });
        });
      });
    
      // Load gists for the user using the GitHub API
      // and build the list.
      function loadGists(user, callback) {
        getUserGists(user, function(gists, error) {
          if (error) {
            $('.gist-list-container').hide();
            $('#error-text').text(JSON.stringify(error, null, 2));
            $('.error-display').show();
    
            if (callback) callback(false);
          } else {
            $('.error-display').hide();
            buildGistList($('#gist-list'), gists, onGistSelected);
            $('.gist-list-container').show();
    
            if (callback) callback(true);
          }
        });
      }
    
      function onGistSelected() {
        $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
        $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
        $('.not-configured-warning').hide();
        $('#settings-done').removeAttr('disabled');
      }
    
      function sendMessage(message) {
        Office.context.ui.messageParent(message);
      }
    
      function getParameterByName(name, url) {
        if (!url) {
          url = window.location.href;
        }
    
        name = name.replace(/[\[\]]/g, "\\$&");
        const regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
          results = regex.exec(url);
    
        if (!results) return null;
    
        if (!results[2]) return '';
    
        return decodeURIComponent(results[2].replace(/\+/g, " "));
      }
    })();
    ```

1. Save your changes.

### Update webpack config settings

Finally, open the **webpack.config.js** file found in the root directory of the project and complete the following steps.

1. Locate the `entry` object within the `config` object and add a new entry for `dialog`.

    ```js
    dialog: "./src/settings/dialog.js",
    ```

    After you've done this, the new `entry` object will look like this:

    ```js
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      dialog: "./src/settings/dialog.js",
    },
    ```

1. Locate the `plugins` array within the `config` object. In the `patterns` array of the `new CopyWebpackPlugin` object, add new entries for **taskpane.css** and **dialog.css**.

    ```js
    {
      from: "./src/taskpane/taskpane.css",
      to: "taskpane.css",
    },
    {
      from: "./src/settings/dialog.css",
      to: "dialog.css",
    },
    ```

    After you've done this, the `new CopyWebpackPlugin` object will look like the following. Note the slight difference if the add-in uses the add-in only manifest.

    ```js
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "./src/taskpane/taskpane.css",
        to: "taskpane.css",
      },
      {
        from: "./src/settings/dialog.css",
        to: "dialog.css",
      },
      {
        from: "assets/*",
        to: "assets/[name][ext][query]",
      },
      {
        from: "manifest*.json", // The file extension is "xml" if the add-in only manifest is being used.
        to: "[name]" + "[ext]",
        transform(content) {
          if (dev) {
            return content;
          } else {
            return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
          }
        },
      },
    ]}),
    ```

1. In the same `plugins` array within the `config` object, add this new object to the end of the array.

    ```js
    new HtmlWebpackPlugin({
      filename: "dialog.html",
      template: "./src/settings/dialog.html",
      chunks: ["polyfill", "dialog"]
    })
    ```

    After you've done this, the new `plugins` array will look ike the following. Note the slight difference if the add-in uses the add-in only manifest.

    ```js
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "./src/taskpane/taskpane.css",
            to: "taskpane.css",
          },
          {
            from: "./src/settings/dialog.css",
            to: "dialog.css",
          },
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json", // The file extension is "xml" if the add-in only manifest is being used.
            to: "[name]." + buildType + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/settings/dialog.html",
        chunks: ["polyfill", "dialog"]
      })
    ],
    ```

### Fetch data from GitHub

The **dialog.js** file you just created specifies that the add-in should load gists when the **change** event fires for the GitHub username field. To retrieve the user's gists from GitHub, you'll use the [GitHub Gists API](https://developer.github.com/v3/gists/).

1. Within the **./src** folder, create a new subfolder named **helpers**.

1. In the **./src/helpers** folder, create a file named **gist-api.js**.

1. In **gist-api.js**, add the following code to retrieve the user's gists from GitHub and build the list of gists.

    ```js
    function getUserGists(user, callback) {
      const requestUrl = 'https://api.github.com/users/' + user + '/gists';
    
      $.ajax({
        url: requestUrl,
        dataType: 'json'
      }).done(function(gists) {
        callback(gists);
      }).fail(function(error) {
        callback(null, error);
      });
    }
    
    function buildGistList(parent, gists, clickFunc) {
      gists.forEach(function(gist) {
    
        const listItem = $('<div/>')
          .appendTo(parent);
    
        const radioItem = $('<input>')
          .addClass('ms-ListItem')
          .addClass('is-selectable')
          .attr('type', 'radio')
          .attr('name', 'gists')
          .attr('tabindex', 0)
          .val(gist.id)
          .appendTo(listItem);
    
        const descPrimary = $('<span/>')
          .addClass('ms-ListItem-primaryText')
          .text(gist.description)
          .appendTo(listItem);
    
        const descSecondary = $('<span/>')
          .addClass('ms-ListItem-secondaryText')
          .text(' - ' + buildFileList(gist.files))
          .appendTo(listItem);
    
        const updated = new Date(gist.updated_at);
    
        const descTertiary = $('<span/>')
          .addClass('ms-ListItem-tertiaryText')
          .text(' - Last updated ' + updated.toLocaleString())
          .appendTo(listItem);
    
        listItem.on('click', clickFunc);
      });  
    }
    
    function buildFileList(files) {
    
      let fileList = '';
    
      for (let file in files) {
        if (files.hasOwnProperty(file)) {
          if (fileList.length > 0) {
            fileList = fileList + ', ';
          }
    
          fileList = fileList + files[file].filename + ' (' + files[file].language + ')';
        }
      }
    
      return fileList;
    }
    ```

1. Save your changes.

1. Run the following command to rebuild the project.

    ```command&nbsp;line
    npm run build
    ```

## Implement a UI-less button

This add-in's **Insert default gist** button is a UI-less button that invokes a JavaScript function, rather than opens a task pane like many add-in buttons do. When the user selects the **Insert default gist** button, the corresponding JavaScript function checks whether the add-in has been configured.

- If the add-in has already been configured, the function loads the content of the gist that the user has selected as the default and inserts it into the body of the message.

- If the add-in hasn't yet been configured, then the settings dialog prompts the user to provide the required information.

### Update the function file (HTML)

A function that's invoked by a UI-less button must be defined in the file that's specified by the `<FunctionFile>` element in the manifest for the corresponding form factor. This add-in's manifest specifies `https://localhost:3000/commands.html` as the function file.

1. Open the **./src/commands/commands.html** and replace the entire contents with the following markup.

    ```html
    <!DOCTYPE html>
    <html>
    
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    
        <!-- Office JavaScript API -->
        <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    
        <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="../../node_modules/showdown/dist/showdown.min.js"></script>
        <script type="text/javascript" src="../../node_modules/urijs/src/URI.min.js"></script>
        <script type="text/javascript" src="../helpers/addin-config.js"></script>
        <script type="text/javascript" src="../helpers/gist-api.js"></script>
    </head>
    
    <body>
      <!-- NOTE: The body is empty on purpose. Since functions in commands.js are
           invoked via a button, there is no UI to render. -->
    </body>
    
    </html>
    ```

    You may have noticed that the HTML file references a JavaScript file, **addin-config.js**, that doesn't yet exist. This file will be created in the [Create a file to manage configuration settings](#create-a-file-to-manage-configuration-settings) section later in this tutorial.

1. Save your changes.

### Update the function file (JavaScript)

1. Open the file **./src/commands/commands.js** and replace the entire contents with the following code. Note that if the **insertDefaultGist** function determines the add-in hasn't yet been configured, it adds the `?warn=1` parameter to the dialog URL. Doing so makes the settings dialog render the message bar that's defined in **./src/settings/dialog.html**, to tell the user why they're seeing the dialog.

    ```js
    let config;
    let btnEvent;
    
    // The onReady function must be run each time a new page is loaded.
    Office.onReady();
    
    function showError(error) {
      Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
        type: 'errorMessage',
        message: error
      });
    }
    
    let settingsDialog;
    
    function insertDefaultGist(event) {
      config = getConfig();
    
      // Check if the add-in has been configured.
      if (config && config.defaultGistId) {
        // Get the default gist content and insert.
        try {
          getGist(config.defaultGistId, function(gist, error) {
            if (gist) {
              buildBodyContent(gist, function (content, error) {
                if (content) {
                  Office.context.mailbox.item.body.setSelectedDataAsync(
                    content,
                    { coercionType: Office.CoercionType.Html },
                    function (result) {
                      event.completed();
                    }
                  );
                } else {
                  showError(error);
                  event.completed();
                }
              });
            } else {
              showError(error);
              event.completed();
            }
          });
        } catch (err) {
          showError(err);
          event.completed();
        }
    
      } else {
        // Save the event object so we can finish up later.
        btnEvent = event;
        // Not configured yet, display settings dialog with
        // warn=1 to display warning.
        const url = new URI('dialog.html?warn=1').absoluteTo(window.location).toString();
        const dialogOptions = { width: 20, height: 40, displayInIframe: true };
    
        Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      }
    }
    
    // Register the function.
    Office.actions.associate("insertDefaultGist", insertDefaultGist);
    
    function receiveMessage(message) {
      config = JSON.parse(message.message);
      setConfig(config, function(result) {
        settingsDialog.close();
        settingsDialog = null;
        btnEvent.completed();
        btnEvent = null;
      });
    }
    
    function dialogClosed(message) {
      settingsDialog = null;
      btnEvent.completed();
      btnEvent = null;
    }
    ```

1. Save your changes.

### Create a file to manage configuration settings

1. In the **./src/helpers** folder, create a file named **addin-config.js** and add the following code. This code uses the [RoamingSettings object](/javascript/api/outlook/office.roamingsettings) to get and set configuration values.

    ```js
    function getConfig() {
      const config = {};
    
      config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
      config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');
    
      return config;
    }
    
    function setConfig(config, callback) {
      Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
      Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);
    
      Office.context.roamingSettings.saveAsync(callback);
    }
    ```

1. Save your changes.

### Create new functions to process gists

1. Open the **./src/helpers/gist-api.js** file and add the following functions. Note the following:

    - If the gist contains HTML, the add-in inserts the HTML as is into the body of the message.

    - If the gist contains Markdown, the add-in uses the [Showdown](https://github.com/showdownjs/showdown) library to convert the Markdown to HTML, then inserts the resulting HTML into the body of the message.

    - If the gist contains anything other than HTML or Markdown, the add-in inserts it into the body of the message as a code snippet.

    ```js
    function getGist(gistId, callback) {
      const requestUrl = 'https://api.github.com/gists/' + gistId;
    
      $.ajax({
        url: requestUrl,
        dataType: 'json'
      }).done(function(gist) {
        callback(gist);
      }).fail(function(error) {
        callback(null, error);
      });
    }
    
    function buildBodyContent(gist, callback) {
      // Find the first non-truncated file in the gist
      // and use it.
      for (let filename in gist.files) {
        if (gist.files.hasOwnProperty(filename)) {
          const file = gist.files[filename];
          if (!file.truncated) {
            // We have a winner.
            switch (file.language) {
              case 'HTML':
                // Insert as is.
                callback(file.content);
                break;
              case 'Markdown':
                // Convert Markdown to HTML.
                const converter = new showdown.Converter();
                const html = converter.makeHtml(file.content);
                callback(html);
                break;
              default:
                // Insert contents as a <code> block.
                let codeBlock = '<pre><code>';
                codeBlock = codeBlock + file.content;
                codeBlock = codeBlock + '</code></pre>';
                callback(codeBlock);
            }
            return;
          }
        }
      }
      callback(null, 'No suitable file found in the gist');
    }
    ```

1. Save your changes.

### Test the Insert default gist button

1. If the local web server isn't already running, run `npm start` from the command prompt.

1. Open Outlook and compose a new message.

1. In the compose message window, select the **Insert default gist** button. You should see a dialog where you can configure the add-in, starting with the prompt to set your GitHub username.

    :::image type="content" source="../images/addin-prompt-configure.png" alt-text="The dialog prompt to configure the add-in.":::

1. In the settings dialog, enter your GitHub username and then either **Tab** or click elsewhere in the dialog to invoke the **change** event, which should load your list of public gists. Select a gist to be the default, and select **Done**.

    :::image type="content" source="../images/addin-settings.png" alt-text="The add-in's settings dialog.":::

1. Select the **Insert default gist** button again. This time, you should see the contents of the gist inserted into the body of the email.

   > [!NOTE]
   > **Classic Outlook on Windows**: To pick up the latest settings, you may need to close and reopen the compose message window.

## Implement a task pane

This add-in's **Display gist list** button opens a task pane and displays the user's gists. The user can then select one of the gists to insert into the body of the message. If the user hasn't yet configured the add-in, they'll be prompted to do so.

### Specify the HTML for the task pane

1. In the project that you've created, the task pane HTML is specified in the file **./src/taskpane/taskpane.html**. Open that file and replace the entire contents with the following markup.

    ```html
    <!DOCTYPE html>
    <html>
    
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Contoso Task Pane Add-in</title>
    
        <!-- Office JavaScript API -->
        <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    
       <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
        <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css"/>
    
        <!-- Template styles -->
        <link href="taskpane.css" rel="stylesheet" type="text/css" />
    </head>
    
    <body class="ms-font-l ms-landing-page">
      <main class="ms-landing-page__main">
        <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
          <div id="not-configured" style="display: none;">
            <div class="centered ms-font-xxl ms-u-textAlignCenter">Welcome!</div>
            <div class="ms-font-xl" id="settings-prompt">Please choose the <strong>Settings</strong> icon at the bottom of this window to configure this add-in.</div>
          </div>
          <div id="gist-list-container" style="display: none;">
            <form>
              <div id="gist-list">
              </div>
            </form>
          </div>
          <div id="error-display" style="display: none;" class="ms-u-borderBase ms-fontColor-error ms-font-m ms-bgColor-error ms-borderColor-error">
          </div>
        </section>
        <button class="ms-Button ms-Button--primary" id="insert-button" tabindex=0 disabled>
          <span class="ms-Button-label">Insert</span>
        </button>
      </main>
      <footer class="ms-landing-page__footer ms-bgColor-themePrimary">
        <div class="ms-landing-page__footer--left">
          <img src="../../assets/logo-filled.png" />
          <h1 class="ms-font-xl ms-fontWeight-semilight ms-fontColor-white">Git the gist</h1>
        </div>
        <div id="settings-icon" class="ms-landing-page__footer--right" aria-label="Settings" tabindex=0>
          <i class="ms-Icon enlarge ms-Icon--Settings ms-fontColor-white"></i>
        </div>
      </footer>
      <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
      <script type="text/javascript" src="../../node_modules/showdown/dist/showdown.min.js"></script>
      <script type="text/javascript" src="../../node_modules/urijs/src/URI.min.js"></script>
      <script type="text/javascript" src="../helpers/addin-config.js"></script>
      <script type="text/javascript" src="../helpers/gist-api.js"></script>
      <script type="text/javascript" src="taskpane.js"></script>
    </body>
    
    </html>
    ```

1. Save your changes.

### Specify the CSS for the task pane

1. In the project that you've created, the task pane CSS is specified in the file **./src/taskpane/taskpane.css**. Open that file and replace the entire contents with the following code.

    ```css
    /* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */
    html, body {
      width: 100%;
      height: 100%;
      margin: 0;
      padding: 0;
      overflow: auto; }
    
    body {
      position: relative;
      font-size: 16px; }
    
    main {
      height: 100%;
      overflow-y: auto; }
    
    footer {
      width: 100%;
      position: relative;
      bottom: 0;
      margin-top: 10px;}
    
    p, h1, h2, h3, h4, h5, h6 {
      margin: 0;
      padding: 0; }
    
    ul {
      padding: 0; }
    
    #settings-prompt {
      margin: 10px 0;
    }
    
    #error-display {
      padding: 10px;
    }
    
    #insert-button {
      margin: 0 10px;
    }
    
    .clearfix {
      display: block;
      clear: both;
      height: 0; }
    
    .pointerCursor {
      cursor: pointer; }
    
    .invisible {
      visibility: hidden; }
    
    .undisplayed {
      display: none; }
    
    .ms-Icon.enlarge {
      position: relative;
      font-size: 20px;
      top: 4px; }
    
    .ms-ListItem-secondaryText,
    .ms-ListItem-tertiaryText {
      padding-left: 15px;
    }
    
    .ms-landing-page {
      display: -webkit-flex;
      display: flex;
      -webkit-flex-direction: column;
              flex-direction: column;
      -webkit-flex-wrap: nowrap;
              flex-wrap: nowrap;
      height: 100%; }
    
    .ms-landing-page__main {
      display: -webkit-flex;
      display: flex;
      -webkit-flex-direction: column;
              flex-direction: column;
      -webkit-flex-wrap: nowrap;
              flex-wrap: nowrap;
      -webkit-flex: 1 1 0;
              flex: 1 1 0;
      height: 100%; }
    
    .ms-landing-page__content {
      display: -webkit-flex;
      display: flex;
      -webkit-flex-direction: column;
              flex-direction: column;
      -webkit-flex-wrap: nowrap;
              flex-wrap: nowrap;
      height: 100%;
      -webkit-flex: 1 1 0;
              flex: 1 1 0;
      padding: 20px; }
    
    .ms-landing-page__content h2 {
      margin-bottom: 20px; }
    
    .ms-landing-page__footer {
      display: -webkit-inline-flex;
      display: inline-flex;
      -webkit-justify-content: center;
              justify-content: center;
      -webkit-align-items: center;
              align-items: center; }
    
    .ms-landing-page__footer--left {
      transition: background ease 0.1s, color ease 0.1s;
      display: -webkit-inline-flex;
      display: inline-flex;
      -webkit-justify-content: flex-start;
              justify-content: flex-start;
      -webkit-align-items: center;
              align-items: center;
      -webkit-flex: 1 0 0px;
              flex: 1 0 0px;
      padding: 20px; }
    
    .ms-landing-page__footer--left:active {
      cursor: default; }
    
    .ms-landing-page__footer--left--disabled {
      opacity: 0.6;
      pointer-events: none;
      cursor: not-allowed; }
    
    .ms-landing-page__footer--left--disabled:active, .ms-landing-page__footer--left--disabled:hover {
      background: transparent; }
    
    .ms-landing-page__footer--left img {
      width: 40px;
      height: 40px; }
    
    .ms-landing-page__footer--left h1 {
      -webkit-flex: 1 0 0px;
              flex: 1 0 0px;
      margin-left: 15px;
      text-align: left;
      width: auto;
      max-width: auto;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis; }
    
    .ms-landing-page__footer--right {
      transition: background ease 0.1s, color ease 0.1s;
      padding: 29px 20px; }
    
    .ms-landing-page__footer--right:active, .ms-landing-page__footer--right:hover {
      background: #005ca4;
      cursor: pointer; }
    
    .ms-landing-page__footer--right:active {
      background: #005ca4; }
    
    .ms-landing-page__footer--right--disabled {
      opacity: 0.6;
      pointer-events: none;
      cursor: not-allowed; }
    
    .ms-landing-page__footer--right--disabled:active, .ms-landing-page__footer--right--disabled:hover {
      background: transparent; }
    ```

1. Save your changes.

### Specify the JavaScript for the task pane

1. In the project that you've created, the task pane JavaScript is specified in the file **./src/taskpane/taskpane.js**. Open that file and replace the entire contents with the following code.

    ```js
    (function() {
      'use strict';
    
      let config;
      let settingsDialog;
    
      Office.onReady(function() {
        $(document).ready(function() {
          config = getConfig();
    
          // Check if add-in is configured.
          if (config && config.gitHubUserName) {
            // If configured, load the gist list.
            loadGists(config.gitHubUserName);
          } else {
            // Not configured yet.
            $('#not-configured').show();
          }
    
          // When insert button is selected, build the content
          // and insert into the body.
          $('#insert-button').on('click', function() {
            const gistId = $('.ms-ListItem.is-selected').val();
            getGist(gistId, function(gist, error) {
              if (gist) {
                buildBodyContent(gist, function (content, error) {
                  if (content) {
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                      content,
                      { coercionType: Office.CoercionType.Html },
                      function (result) {
                        if (result.status === Office.AsyncResultStatus.Failed) {
                          showError("Could not insert gist: " + result.error.message);
                        }
                      }
                    );
                  } else {
                    showError('Could not create insertable content: ' + error);
                  }
                });
              } else {
                showError('Could not retrieve gist: ' + error);
              }
            });
          });
    
          // When the settings icon is selected, open the settings dialog.
          $('#settings-icon').on('click', function() {
            // Display settings dialog.
            let url = new URI('dialog.html').absoluteTo(window.location).toString();
            if (config) {
              // If the add-in has already been configured, pass the existing values
              // to the dialog.
              url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
            }
    
            const dialogOptions = { width: 20, height: 40, displayInIframe: true };
    
            Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
              settingsDialog = result.value;
              settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
              settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
            });
          })
        });
      });
    
      function loadGists(user) {
        $('#error-display').hide();
        $('#not-configured').hide();
        $('#gist-list-container').show();
    
        getUserGists(user, function(gists, error) {
          if (error) {
    
          } else {
            $('#gist-list').empty();
            buildGistList($('#gist-list'), gists, onGistSelected);
          }
        });
      }
    
      function onGistSelected() {
        $('#insert-button').removeAttr('disabled');
        $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
        $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
      }
    
      function showError(error) {
        $('#not-configured').hide();
        $('#gist-list-container').hide();
        $('#error-display').text(error);
        $('#error-display').show();
      }
    
      function receiveMessage(message) {
        config = JSON.parse(message.message);
        setConfig(config, function(result) {
          settingsDialog.close();
          settingsDialog = null;
          loadGists(config.gitHubUserName);
        });
      }
    
      function dialogClosed(message) {
        settingsDialog = null;
      }
    })();
    ```

1. Save your changes.

### Test the "Display gist list" button

1. If the local web server isn't already running, run `npm start` from the command prompt.

1. Open Outlook and compose a new message.

1. In the compose message window, select the **Display gist list** button. You should see a task pane open to the right of the compose form.

1. In the task pane, select the **Hello World Html** gist and select **Insert** to insert that gist into the body of the message.

:::image type="content" source="../images/addin-taskpane.png" alt-text="The add-in task pane and the selected gist content displayed in the message body.":::

## Next steps

In this tutorial, you've created an Outlook add-in that can be used in message compose mode to insert content into the body of a message. To learn more about developing Outlook add-ins, continue to the following article.

> [!div class="nextstepaction"]
> [Outlook add-in APIs](../outlook/apis.md)

## Code samples

- [Completed Outlook add-in tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/outlook-tutorial): The result of completing this tutorial.

## See also

- [Office Add-in manifests](../develop/add-in-manifests.md)
- [Outlook add-in design guidelines](../outlook/outlook-addin-design.md)
- [Debug function commands in Outlook add-ins](../outlook/debug-ui-less.md)
