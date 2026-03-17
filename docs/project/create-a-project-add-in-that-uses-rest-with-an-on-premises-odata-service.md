---
title: Create a Project add-in that uses REST with an on-premises Project Server OData service
description: Learn how to build a task pane add-in for Project Professional that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance.
ms.date: 03/17/2026
ms.localizationpriority: medium
---

# Create a Project add-in that uses REST with an on-premises Project Server OData service

This article describes how to build a task pane add-in for Project Professional that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance. The add-in uses REST to access the **ProjectData** OData reporting service in Project Server.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project Professional 2016 or later on Windows. You need Project Professional to connect with Project Web App.

    > [!NOTE]
    > Project Standard can also host task pane add-ins, but can't sign in to Project Web App.

- Access to a Project Web App instance in an on-premises installation of Project Server. The procedures and code examples in this article access the **ProjectData** service of Project Server in a local domain.

### Verify that you can access the ProjectData service

1. Query the **ProjectData** service by using your browser with the following URL: `http://ServerName/ProjectServerName/_api/ProjectData`. For example, if the Project Web App instance is `http://MyServer/pwa`, browse to `http://MyServer/pwa/_api/ProjectData`.

    The browser should show XML results similar to the following.

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

1. You might need to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you don't have sign-in permission for that Project Web App instance, or there's a network problem that requires administrative help.

## Create the add-in project

Because the Yeoman generator for Office Add-ins doesn't have a dedicated Project task pane template with full scaffolding, use the manifest-only option and then create the web application files manually.

1. Run the following command to create an add-in project by using the Yeoman generator.

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > When you run the `yo office` command, you might receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

    When prompted, provide the following information to create your add-in project.

    - **Choose a project type:** `Office Add-in project containing the manifest only`
    - **What do you want to name your add-in?** `HelloProjectOData`
    - **Which Office client application would you like to support?** `Project`

1. Go to the project folder.

    ```command&nbsp;line
    cd HelloProjectOData
    ```

## Set up the project structure

The manifest-only project contains a `manifest.xml` file. You need to create the web application files that provide the add-in UI and functionality.

### Create the file structure

1. Create the following folder structure in the project root.

    ```command&nbsp;line
    mkdir src
    mkdir src\taskpane
    ```

1. Create the following files in the `src\taskpane` folder. The following sections provide the contents for each file:

    - `taskpane.html` - The HTML markup for the task pane.
    - `taskpane.css` - The CSS styles for the task pane.
    - `taskpane.js` - The JavaScript code that interacts with the Office application and the OData service.

### Update the manifest

Open the `manifest.xml` file and make the following changes.

1. Change the `<Description>` value to `Compares cost and work data in the active project with averages for all projects`.
1. Verify that the `<SourceLocation>` element points to your task pane HTML file. Update it to:

    ```xml
    <SourceLocation DefaultValue="https://localhost:3000/src/taskpane/taskpane.html" />
    ```

### Create the HTML content

Create the `src\taskpane\taskpane.html` file with the following content. The task pane provides two buttons and a comparison table:

- **Get ProjectData Endpoint** gets the OData service URL from the active Project Web App connection.
- **Compare All Projects** queries the OData service and displays average values alongside the current project values.

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>
    <link rel="stylesheet" type="text/css" href="taskpane.css" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script src="taskpane.js"></script>
</head>
<body>
    <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the
                <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
                onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
    </div>
    <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
                <td class="heading_leftCol"></td>
                <td class="heading_midCol"><strong>Average</strong></td>
                <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
                <td class="row_leftCol"><strong>Project Cost</strong></td>
                <td class="row_midCol" id="AverageProjectCost">&nbsp;</td>
                <td class="row_rightCol" id="CurrentProjectCost">&nbsp;</td>
            </tr>
            <tr>
                <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                <td class="row_midCol" id="AverageProjectActualCost">&nbsp;</td>
                <td class="row_rightCol" id="CurrentProjectActualCost">&nbsp;</td>
            </tr>
            <tr>
                <td class="row_leftCol"><strong>Project Work</strong></td>
                <td class="row_midCol" id="AverageProjectWork">&nbsp;</td>
                <td class="row_rightCol" id="CurrentProjectWork">&nbsp;</td>
            </tr>
            <tr>
                <td class="row_leftCol"><strong>Project % Complete</strong></td>
                <td class="row_midCol" id="AverageProjectPercentComplete">&nbsp;</td>
                <td class="row_rightCol" id="CurrentProjectPercentComplete">&nbsp;</td>
            </tr>
        </table>
    </div>
    <br />
    <textarea id="odataText" rows="12" cols="40"></textarea>
</body>
</html>
```

### Create the CSS styles

Create the file `src\taskpane\taskpane.css` with the following content.

```css
body {
    font-size: 11pt;
}

h1 {
    font-size: 22pt;
}

h2 {
    font-size: 16pt;
}

.rest {
    font-family: 'Courier New';
    font-size: 0.9em;
}

.button-wide {
    width: 210px;
    margin-top: 2px;
}

.button-narrow {
    width: 80px;
    margin-top: 2px;
}

.infoTable {
    text-align: center;
    vertical-align: middle;
}

.heading_leftCol {
    width: 20px;
    height: 20px;
}

.heading_midCol {
    width: 100px;
    height: 20px;
    font-size: medium;
    font-weight: bold;
}

.heading_rightCol {
    width: 101px;
    height: 20px;
    font-size: medium;
    font-weight: bold;
}

.row_leftCol {
    width: 20px;
    font-size: small;
    font-weight: bold;
}

.row_midCol {
    width: 100px;
}

.row_rightCol {
    width: 101px;
}
```

### Create the JavaScript code

Create the file `src\taskpane\taskpane.js` with the following content. The code is explained in more detail below.

```js
const PROJDATA = "/_api/ProjectData";
const PROJQUERY = "/Projects?";
const QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
const QUERY_SELECT1 = "&$select=ProjectId, ProjectName";
const QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
let _pwa;           // URL of Project Web App.
let _projectUid;    // GUID of the active project.
let _docUrl;        // Path of the project document.
let _odataUrl = ""; // URL of the OData service: http[s]://ServerName/ProjectServerName/_api/ProjectData

// Ensure the Office.js library is loaded.
Office.onReady(function () {
    // Office is ready.
});

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project isn't connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                if (_pwa.substring(0, 4) === "http") {
                    _odataUrl = _pwa + PROJDATA;
                    document.getElementById("compareProjects").removeAttribute("disabled");
                    getProjectGuid();
                } else {
                    _odataUrl = "No connection!";
                    showError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                document.getElementById("projectDataEndPoint").textContent = _odataUrl;
            } else {
                showError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            } else {
                showError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project Web App, which is in the form <>\ProjectName.
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

// Get data about all projects on Project Server,
// by using a REST query with the fetch API.
function retrieveOData() {
    const restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;

    fetch(restUrl, {
        method: "GET",
        headers: {
            "Accept": "application/json; odata=verbose"
        },
        credentials: "include"
    })
    .then(function (response) {
        const contentType = response.headers.get("Content-Type");
        const status = response.status;
        return response.text().then(function (responseText) {
            // Create a message to display in the text box.
            const message = "\r\ntextStatus: " + (response.ok ? "success" : "error") +
                "\r\nContentType: " + contentType +
                "\r\nStatus: " + status +
                "\r\nResponseText:\r\n" + responseText;

            // Parse and display the JSON response.
            parseODataResult(responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            const odataText = document.getElementById("odataText");
            odataText.textContent = _docUrl;
            odataText.textContent += "\r\nREST query:\r\n" + restUrl;
            odataText.textContent += message;
        });
    })
    .catch(function (error) {
        const odataText = document.getElementById("odataText");
        odataText.textContent = "Error: " + error.message;
        showError("Network error", error.message);
    });
}

// Calculate the average values of actual cost, cost, work, and percent complete
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    const res = JSON.parse(oDataResult);
    const len = res.d.results.length;
    let projActualCost = 0;
    let projCost = 0;
    let projWork = 0;
    let projPercentCompleted = 0;
    let myProjectIndex = -1;

    for (let i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,
        // store the project index.
        if (currentProjectGuid.toLocaleLowerCase() === res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
    }

    const avgProjCost = (projCost / len).toFixed(2);
    const avgProjWork = (projWork / len).toFixed(1);
    const avgProjActualCost = (projActualCost / len).toFixed(2);
    const avgProjPercentCompleted = (projPercentCompleted / len).toFixed(1);

    // Display averages in the table, with the correct units.
    document.getElementById("AverageProjectCost").textContent = "$" + avgProjCost;
    document.getElementById("AverageProjectActualCost").textContent = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").textContent = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").textContent = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex !== -1) {
        const myProjCost = Number(res.d.results[myProjectIndex].ProjectCost).toFixed(2);
        const myProjWork = Number(res.d.results[myProjectIndex].ProjectWork).toFixed(1);
        const myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost).toFixed(2);
        const myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted).toFixed(1);

        setComparisonValue("CurrentProjectCost", "$" + myProjCost, Number(myProjCost) <= Number(avgProjCost));
        setComparisonValue("CurrentProjectActualCost", "$" + myProjActualCost, Number(myProjActualCost) <= Number(avgProjActualCost));
        setComparisonValue("CurrentProjectWork", myProjWork + " hrs", Number(myProjWork) > Number(avgProjWork));
        setComparisonValue("CurrentProjectPercentComplete", myProjPercentCompleted + "%", Number(myProjPercentCompleted) > Number(avgProjPercentCompleted));
    } else {
        // The current project isn't published.
        const naFields = ["CurrentProjectCost", "CurrentProjectActualCost", "CurrentProjectWork", "CurrentProjectPercentComplete"];
        naFields.forEach(function (id) {
            document.getElementById(id).textContent = "NA";
            document.getElementById(id).style.color = "blue";
        });
    }
}

// Helper function to set a comparison value with color coding.
// Green means favorable, red means unfavorable.
function setComparisonValue(elementId, text, isFavorable) {
    const element = document.getElementById(elementId);
    element.textContent = text;
    element.style.color = isFavorable ? "green" : "red";
}

// Display an error message in a notification area.
function showError(title, message) {
    const odataText = document.getElementById("odataText");
    odataText.textContent = "Error: " + title + "\r\n" + message;
}
```

## Understand the code

The JavaScript includes global constants for the REST query and global variables that several functions use. Here's how the key functions work.

### setOdataUrl

The **Get ProjectData Endpoint** button calls `setOdataUrl`, which uses the [getProjectFieldAsync method](/javascript/api/office/office.document) to get the Project Web App URL. If Project is connected with Project Web App, the function enables the **Compare All Projects** button and displays the **ProjectData** service URL. If Project isn't connected, the function displays an error message.

### retrieveOData

When the user selects **Compare All Projects**, the `retrieveOData` function builds a REST query URL and calls the **ProjectData** OData service by using the Fetch API. The REST query filters out administrative projects and selects cost, work, and percent complete fields.

> [!NOTE]
> This code works with an on-premises installation of Project Server. For Project on the web, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).

### parseODataResult

The `parseODataResult` function calculates average values of cost and work data across all projects, then compares them with the current project. It color-codes the values:

- **Green**: The current project value is favorable (lower cost or higher work/completion).
- **Red**: The current project value is unfavorable.
- **Blue NA**: The current project isn't published to Project Server.

## Serve the add-in locally

To serve your add-in files, you need a local web server. You can use any HTTP server you prefer. The following steps use the `http-server` npm package as an example.

1. Install a development server.

    ```command&nbsp;line
    npm install --save-dev http-server
    ```

1. Generate development certificates for HTTPS.

    ```command&nbsp;line
    npx office-addin-dev-certs install
    ```

    Accept the prompt to install the certificate. The process stores the certificates in your user profile directory at `%USERPROFILE%\.office-addin-dev-certs\`.

1. Add the following script to the `package.json` file's `"scripts"` section. This script references the certificate location where `office-addin-dev-certs` stores them.

    ```json
    "start": "http-server . --ssl --cert \"%USERPROFILE%\\.office-addin-dev-certs\\localhost.crt\" --key \"%USERPROFILE%\\.office-addin-dev-certs\\localhost.key\" -p 3000"
    ```

1. Start the local server.

    ```command&nbsp;line
    npm start
    ```

## Test the add-in

To test the **HelloProjectOData** add-in, you must install Project Professional on your development computer. To enable different test scenarios, make sure you can choose whether Project opens for files on the local computer or connects with Project Web App.

1. In Project Professional, on the **File** tab, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.

1. In the **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account. In the **When starting** section, select **Choose an account**.

### Sideload the add-in

1. Follow the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to sideload the add-in in Project by using the manifest file.

### Run through test scenarios

1. **Test with a published project**: Connect with Project Web App and open a published project that contains cost and work data. On the **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData**. Select **Get ProjectData Endpoint**, and then select **Compare All Projects**. Verify that the add-in displays the endpoint and correctly displays the cost and work data in the comparison table.

1. **Test without a Project Web App connection**: Open a local .mpp file without connecting to Project Web App. Open the **Hello ProjectData** task pane and select **Get ProjectData Endpoint**. The add-in should show a "No connection!" error, and the **Compare All Projects** button should remain disabled.

1. **Test with an unpublished project**: Connect to Project Web App and create a project with cost and work data. Save the project but don't publish it. Open the **Hello ProjectData** task pane and compare projects. You should see a blue **NA** for fields in the **Current** column.

> [!NOTE]
> There are limits to the amount of data that one query of the **ProjectData** service can return. The amount of data varies by entity. For example, the `Projects` entity set has a default limit of 100 projects per query. For a production add-in, modify the code to enable queries of more than 100 projects. For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).

## Next steps

If **HelloProjectOData** were a production add-in, you'd design it differently. For example, you wouldn't include debug output in a text box, and you probably wouldn't add a button to get the **ProjectData** endpoint. You'd also need to rewrite the `retrieveOData` function to handle Project Web App instances that have more than 100 projects.

The add-in should include additional error checks and logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1,000 projects with an average duration of five days and average cost of $2,400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. You could show that with a frequency graph. You might add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.

For other queries of the **ProjectData** service, query string length limits affect the number of steps that a query can take from a parent collection to an object in a child collection. For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item might exceed the default maximum URL length. For more information, see [Query OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).

For production use, consider the following improvements.

- Rewrite the `retrieveOData` function to enable queries of more than 100 projects. For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query and use the *$skip* operator and *$top* operator in the REST query for project data. Run multiple queries in a loop and then average the data from each query. Each query for project data would be of the form:

  `~/ProjectData/Projects()?skip=[numSkipped]&$top=100&$filter=[filter]&$select=[field1,field2, ...]`

  For more information, see [OData system query options using the REST endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/microsoft.sharepoint.powershell/set-spprojectodataconfiguration) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).

- To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).

## See also

- [Task pane add-ins for Project](project-add-ins.md)
- [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15))
- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Publish your Office Add-in](../publish/publish.md)
