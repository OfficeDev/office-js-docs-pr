### Configuration

The following files specify configuration settings for the add-in.

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.

- The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.
    > [!NOTE]
    > Some of the constants defined in this file are used to facilitate the SSO process. 

### Task pane 

The following files define the add-in's task pane UI and functionality.

- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.

### Authentication

The following files facilitate the SSO process and write data to the Office document.

- The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.
- The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.
- The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.
- The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.
- The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.