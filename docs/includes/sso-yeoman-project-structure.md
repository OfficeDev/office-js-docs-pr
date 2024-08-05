### Configuration

The following files specify configuration settings for the add-in.

- The **./manifest.xml** or **manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.

- The **./.ENV** file in the root directory of the project defines constants that are used by the add-in project.

### Task pane

The following files define the add-in's task pane UI and functionality.

- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.

- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.

- The **./src/taskpane/taskpane.js** file contains code to initialize the add-in and also code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document.

### Authentication

The following files facilitate the SSO process and write data to the Office document.

- In a JavaScript project, the **./src/helpers/documentHelper.js** file contains code that encapsulates the user's profile information for insertion into the current Office document. There's no such file in a TypeScript project. Instead, the code that gathers the profile information is inline in the **./src/taskpane/taskpane.ts** file.

- The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the JavaScript for the fallback authentication strategy. The `<script>` tag to load the JavaScript is inserted into the file when Webpack.config.js runs.

- The **./src/helpers/fallbackauthdialog.js** file contains the JavaScript for the fallback authentication strategy that signs in the user with msal.js.

- The **./src/helpers/message-helper.js** file contains JavaScript that shows or hides error messages to the user.

- The **./src/helpers/middle-tier-calls.js** file contains the JavaScript that calls your web API for fetching data.

- The **./src/helpers/sso-helper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the access token, and includes it in a call to Microsoft Graph for the data. In the event of an error or in scenarios when SSO authentication isn't supported, it invokes the fallback strategy. 
