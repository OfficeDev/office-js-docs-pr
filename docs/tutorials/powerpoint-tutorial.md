---
title: PowerPoint add-in tutorial
description: In this tutorial, you will build an PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.
ms.date: 03/31/2023
ms.prod: powerpoint
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
> - Navigates between slides

## Prerequisites

- [Visual Studio 2019, *version 16.10.3 or earlier*, or Visual Studio 2022](https://www.visualstudio.com/vs/), with the **Office/SharePoint development** workload installed.

    > [!IMPORTANT]
    > Some versions of Visual Studio 2019 after 16.10.3 have a bug that prevents this tutorial from being completed. Use an earlier version of Visual Studio 2019 or use Visual Studio 2022.

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

1. The following NuGet packages must be installed. Install them using the **NuGet Package Manager** in Visual Studio. See Visual Studio help for instructions. The second of these may be installed automatically when you install the first.

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
            <!-- TODO5: Create the go-to-slide buttons. -->
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
                // Initialize the FabricUI notification mechanism and hide it
                const element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
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

1. Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.

1. In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.

1. In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button. Visual Studio creates and opens the **PhotoController.cs** file.

    > [!NOTE]
    > The scaffolding process does not complete properly on some versions of Visual Studio 2019 after version 16.10.3. Visual Studio 2022 is not affected.

1. Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.

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

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
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
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
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

1. Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.

    ![The PowerPoint add-in with the Insert Image button highlighted.](../images/powerpoint-tutorial-insert-image-button.png)

1. In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

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

1. Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![The Show Taskpane button highlighted on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. Notice that the task pane now contains a header section and title, and no longer contains a footer section.

    ![The PowerPoint add-in with Insert Image button.](../images/powerpoint-tutorial-new-task-pane-ui.png)

1. In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

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

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![The Show Taskpane button on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.

    ![The selected PowerPoint title slide and the Insert Image button highlighted in the add-in.](../images/powerpoint-tutorial-insert-image-slide-design.png)

1. Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.

    ![The selected PowerPoint title slide with the Insert Text button highlighted in the add-in.](../images/powerpoint-tutorial-insert-text.png)

1. In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

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

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![The Show Taskpane button on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.

    ![The Get Slide Metadata button highlighted in the add-in.](../images/powerpoint-tutorial-get-slide-metadata.png)

1. In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Navigate between slides

Complete the following steps to add code that navigates between the slides of a document.

1. In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.

    ```html
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

1. In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

1. In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions. Each of these functions uses the `goToByIdAsync` method to select a slide based upon its position in the document (first, last, previous, and next).

    ```js
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

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted on the Visual Studio toolbar.](../images/powerpoint-tutorial-start.png)

1. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.

1. In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.

    ![The Go to First Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-first-slide.png)

1. In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.

    ![The Go to Next Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-next-slide.png)

1. In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.

    ![The Go to Previous Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-previous-slide.png)

1. In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.

    ![The Go to Last Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-last-slide.png)

1. In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted on the Visual Studio toolbar.](../images/powerpoint-tutorial-stop.png)

## Next steps

In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides. To learn more about building PowerPoint add-ins, continue to the following article.

> [!div class="nextstepaction"]
> [PowerPoint add-ins overview](../powerpoint/powerpoint-add-ins.md)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
