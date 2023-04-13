### Configuration

The following files specify configuration settings for the add-in.

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.

- The **./.ENV** file in the root directory of the project defines constants that are used by the add-in project.

### Task pane

The following files define the add-in's task pane UI and functionality.

- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.

- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.

- In a JavaScript project, the **./src/taskpane/taskpane.js** file contains code to initialize the add-in. In a TypeScript project, the **./src/taskpane/taskpane.ts** file contains code to initialize the add-in and also code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document.

### Authentication

The following files facilitate the SSO process and write data to the Office document.

- In a JavaScript project, the **./src/helpers/documentHelper.js** file contains code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document. There is no such file in a TypeScript project; the code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document exists in **./src/taskpane/taskpane.ts** instead.

- The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the JavaScript for the fallback authentication strategy.

- The **./src/helpers/fallbackauthdialog.js** file contains the JavaScript for the fallback authentication strategy that signs in the user with msal.js.

- The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication strategy in scenarios when SSO authentication is not supported.

- The **./src/middle-tier/ssoauth-helper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the access token, initiates the swap of the access token for a new access token with permissions to Microsoft Graph, and calls to Microsoft Graph for the data.
