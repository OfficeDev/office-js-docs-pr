---
title: PowerPoint add-in tutorial
description: In this tutorial, you will build an PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.
ms.date: 09/19/2023
ms.service: powerpoint
#Customer intent: As a developer, I want to build a PowerPoint add-in that can interact with content in a PowerPoint document.
ms.localizationpriority: high
---

# Tutorial: Create a PowerPoint task pane add-in

In this tutorial, you'll use Visual Studio to create an PowerPoint task pane add-in that:

> [!div class="checklist"]
>
> - Adds the [Bing](https://www.bing.com) photo of the day to a slide
> - Adds text to a slide
> - Gets slide metadata
> - Adds new slides
> - Navigates between slides

## Create the add-in

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# [Yeoman generator](#tab/yeomangenerator)

> [!TIP]
> If you've already completed the [Build your first PowerPoint task pane add-in](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) quick start using the Yeoman generator, and want to use that project as a starting point for this tutorial, go directly to the [Insert an image](#insert-an-image) section to start this tutorial.

### Prerequisites

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### Create the add-in project

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `JavaScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `PowerPoint`

![Screenshot showing the prompts and answers for the Yeoman generator in a command line interface.](../images/yo-office-powerpoint.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### Complete setup

1. Navigate to the root directory of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. This add-in will use the following library.

    - [jQuery](https://jquery.com/) library to simplify DOM interactions.

     To install these tools for your project, run the following command in the root directory of the project.

    ```command&nbsp;line
    npm install jquery --save
    ```

1. Open your project in VS Code or your preferred code editor.

    [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

### Insert an image

Complete the following steps to add code that retrieves [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.

1. Open the project in your code editor.

1. Open the file **./src/taskpane/taskpane.html**. This file contains the HTML markup for the task pane.

1. Locate the `<main>` element and replace it with the following markup, and save the file.

    ```html
    <!-- TODO: Need to provide instructions for this!! -->
    <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>

    <!-- TODO2: Update the header node. -->
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the add-slides and go-to-slide buttons. -->
        </div>
    </main>
    ```

1. Open the file **./src/taskpane/taskpane.js**. This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application. Replace the entire contents with the following code and save the file.

    ```js
    /*
     * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
     * See LICENSE in the project root for license information.
     */
    
    /* global document, Office */
    
    Office.onReady((info) => {
      if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        // TODO1: Assign event handler for insert-image button.
        // TODO4: Assign event handler for insert-text button.
        // TODO6: Assign event handler for get-slide-metadata button.
        // TODO8: Assign event handlers for add-slides and the four navigation buttons.
      }
    });
    
    // TODO2: Define the insertImage function.

    // TODO3: Define the insertImageFromBase64String function.

    // TODO5: Define the insertText function.

    // TODO7: Define the getSlideMetadata function.

    // TODO9: Define the addSlides and navigation functions.

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            console.error(error);
        }
    }
    ```

1. In the **taskpane.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.

    ```html
    <button class="ms-Button" id="insert-image">Insert Image</button>
    ```

1. In the **taskpane.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.

    ```js
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    ```

1. In the **taskpane.js** file, replace `TODO2` with the following code to define the `insertImage` function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.

    ```js
    function insertImage() {
        // Get image from web service (as a Base64-encoded string).
        const bingUrl = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1";
        $.ajax({
            url: bingUrl,
            dataType: "json",
            success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

1. In the **taskpane.js** file, replace `TODO3` with the following code to define `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:

    - The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsync` request indicates the type of data being inserted.

    - The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. In PowerPoint, insert a new blank slide, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.

    ![Screenshot of PowerPoint with the Show Taskpane button highlighted.](../images/powerpoint_quickstart_addin_1c.png)

1. At the bottom of the task pane, choose the **Run** link to insert the text "Hello World" into the current slide.

    ![Screenshot of PowerPoint with an image of a dog and the text 'Hello World` displayed on the slide.](../images/powerpoint_quickstart_addin_3c.png)

# [Visual Studio](#tab/visualstudio)

> [!TIP]
> If you want a completed version of this tutorial, head over to the [Office Add-ins samples repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial).

## Prerequisites

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/), with the **Office/SharePoint development** workload installed.

    > [!NOTE]
    > If you've previously installed Visual Studio, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).

    > [!NOTE]
    > If you don't already have Office, you can [join the Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program) to get a free, 90-day renewable Microsoft 365 subscription to use during development.

## Create your add-in project

Complete the following steps to create a PowerPoint add-in project using Visual Studio.

1. Choose **Create a new project**.

1. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.

1. Name the project `HelloWorld`, and select **Create**.

1. In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.

1. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

     ![The Visual Studio Solution Explorer window showing HelloWorld and HelloWorldWeb, the two projects in the HelloWorld solution.](../images/powerpoint-tutorial-solution-explorer.png)

1. The following NuGet packages must be installed. Install them on the **HelloWorldWeb** project using the **NuGet Package Manager** in Visual Studio. See Visual Studio help for instructions. The second of these may be installed automatically when you install the first.

   - Microsoft.AspNet.WebApi.WebHost
   - Microsoft.AspNet.WebApi.Core

   > [!IMPORTANT]
   > When you're using the **NuGet Package Manager** to install these packages, do **not** install the recommended update to jQuery. The jQuery version installed with your Visual Studio solution matches the jQuery call within the solution files.

### Explore the Visual Studio solution

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### Update code

Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.

1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the add-slides and go-to-slide buttons. -->
        </div>
    </div>
    ```

1. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.

    ```js
    (function () {
        "use strict";

        let messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it.
                const element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for add-slides and the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function.

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the addSlides and navigation functions.

        // Helper function for displaying notifications.
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```

## Insert an image

Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.

1. Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.

    ![The Visual Studio Solution Explorer window showing the Controllers folder highlighted in the HelloWorldWeb project.](../images/powerpoint-tutorial-solution-explorer-controllers.png)

1. Right-click the **Controllers** folder and select **Add** > **New Scaffolded Item...**.

1. In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.

1. In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button. Visual Studio creates and opens the **PhotoController.cs** file.

    > [!IMPORTANT]
    > The scaffolding process doesn't complete properly on some versions of Visual Studio after version 16.10.3. If you have the **Global.asax** and **./App_Start/WebApiConfig.cs** files, then skip to step 6.
    >
    > ![The Visual Studio Solution Explorer window showing the scaffolded files highlighted in the HelloWorldWeb project.](../images/powerpoint-tutorial-solution-explorer-scaffolded.png)

1. If you're missing scaffolding files from the **HelloWorldWeb** project, add them as follows.

    1. Using Solution Explorer, add a new folder named **App_Start** to the **HelloWorldWeb** project.

    1. Right-click the **App_Start** folder and select **Add** > **Class...**.

    1. In the **Add New Item** dialog, name the file **WebApiConfig.cs** then choose the **Add** button.

    1. Replace the entire contents of the **WebApiConfig.cs** file with the following code.

        ```cs
        using System;
        using System.Collections.Generic;
        using System.Linq;
        using System.Web;
        using System.Web.Http;
        
        namespace HelloWorldWeb.App_Start
        {
            public static class WebApiConfig
            {
                public static void Register(HttpConfiguration config)
                {
                    config.MapHttpAttributeRoutes();
        
                    config.Routes.MapHttpRoute(
                        name: "DefaultApi",
                        routeTemplate: "api/{controller}/{id}",
                        defaults: new { id = RouteParameter.Optional }
                    );
                }
            }
        }
        ```

    1. In the Solution Explorer, right-click the **HelloWorldWeb** project and select **Add** > **New Item...**.

    1. In the **Add New Item** dialog, search for "global", select **Global Application Class**, then choose the **Add** button. By default, the file is named **Global.asax**.

    1. Replace the entire contents of the **Global.asax.cs** file with the following code.

        ```cs
        using HelloWorldWeb.App_Start;
        using System;
        using System.Collections.Generic;
        using System.Linq;
        using System.Web;
        using System.Web.Http;
        using System.Web.Security;
        using System.Web.SessionState;
        
        namespace HelloWorldWeb
        {
            public class WebApiApplication : System.Web.HttpApplication
            {
                protected void Application_Start()
                {
                    GlobalConfiguration.Configure(WebApiConfig.Register);
                }
            }
        }
        ```

    1. In the Solution Explorer, right-click the **Global.asax** file and choose **View Markup**.

    1. Replace the entire contents of the **Global.asax** file with the following code.

        ```XML
        <%@ Application Codebehind="Global.asax.cs" Inherits="HelloWorldWeb.WebApiApplication" Language="C#" %>
        ```

1. Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64-encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64-encoded string.

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the XML response and get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64-encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

1. In the **Home.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

1. In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.

    ```js
    $('#insert-image').click(insertImage);
    ```

1. In the **Home.js** file, replace `TODO2` with the following code to define the `insertImage` function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.

    ```js
    function insertImage() {
        // Get image from web service (as a Base64-encoded string).
        $.ajax({
            url: "/api/photo/",
            dataType: "text",
            success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

1. In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:

    - The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsync` request indicates the type of data being inserted.

    - The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.

    ![The PowerPoint add-in with the Insert Image button highlighted.](../images/powerpoint-tutorial-insert-image-button.png)

    > [!NOTE]
    > If you get an error "Could not find file [...]\bin\roslyn\csc.exe", then do the following:
    >
    > 1. Open the **.\Web.config** file.
    > 1. Find the **\<compiler\>** node for the .cs `extension`, then remove the `type` attribute and its value.
    > 1. Save the file.

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Customize User Interface (UI) elements

Complete the following steps to add markup that customizes the task pane UI.

1. In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane. Note:

    - The styles that begin with `ms-` are defined by [Fabric Core in Office Add-ins](../design/fabric-core.md), a JavaScript front-end framework for building user experiences for Office. The **Home.html** file includes a reference to the Fabric Core stylesheet.

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

1. In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.

### Test the add-in

1. Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. Notice that the task pane now contains a header section and title, and no longer contains a footer section.

    ![The PowerPoint add-in with Insert Image button.](../images/powerpoint-tutorial-new-task-pane-ui.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Insert text

Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.

1. In the **Home.html** file, replace `TODO3` with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

1. In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.

    ```js
    $('#insert-text').click(insertText);
    ```

1. In the **Home.js** file, replace `TODO5` with the following code to define the `insertText` function. This function inserts text into the current slide.

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.

    ![The selected PowerPoint title slide and the Insert Image button highlighted in the add-in.](../images/powerpoint-tutorial-insert-image-slide-design.png)

1. Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.

    ![The selected PowerPoint title slide with the Insert Text button highlighted in the add-in.](../images/powerpoint-tutorial-insert-text.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Get slide metadata

Complete the following steps to add code that retrieves metadata for the selected slide.

1. In the **Home.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

1. In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

1. In the **Home.js** file, replace `TODO7` with the following code to define the `getSlideMetadata` function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### Test the add-in

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.

    ![The Get Slide Metadata button highlighted in the add-in.](../images/powerpoint-tutorial-get-slide-metadata.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Navigate between slides

Complete the following steps to add code that navigates between the slides of a document.

1. In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="add-slides">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Add Slides</span>
        <span class="Button-description">Adds 2 slides.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
    </button>
    ```

1. In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the **Add Slides** and four navigation buttons.

    ```js
    $('#add-slides').click(addSlides);
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

1. In the **Home.js** file, replace `TODO9` with the following code to define the `addSlides` and navigation functions. Each of these functions uses the `goToByIdAsync` method to select a slide based upon its position in the document (first, last, previous, and next).

    ```js
    async function addSlides() {
        await PowerPoint.run(async function (context) {
            context.presentation.slides.add();
            context.presentation.slides.add();

            await context.sync();

            showNotification("Success", "Slides added.");
            goToLastSlide();
        });
    }

    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted on the Visual Studio toolbar.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Add Slides** button. Two new slides are added to the document and the last slide in the document is selected and displayed.

    ![The Add Slides button highlighted in the add-in.](../images/powerpoint-tutorial-add-slides.png)

1. In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.

    ![The Go to First Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-first-slide-1.png)

1. In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.

    ![The Go to Next Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-next-slide-1.png)

1. In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.

    ![The Go to Previous Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-previous-slide-1.png)

1. In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.

    ![The Go to Last Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-last-slide-1.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted on the Visual Studio toolbar.](../images/powerpoint-tutorial-stop.png)

---

## Next steps

In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides. To learn more about building PowerPoint add-ins, continue to the following article.

> [!div class="nextstepaction"]
> [PowerPoint add-ins overview](../powerpoint/powerpoint-add-ins.md)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
