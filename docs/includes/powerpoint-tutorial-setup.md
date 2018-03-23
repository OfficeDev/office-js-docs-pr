You'll begin this tutorial by setting up your development project. 

> [!NOTE]
> This page describes an individual step of the PowerPoint add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.

## Prerequisites

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

## Setup

In this tutorial, you'll create an add-in using Visual Studio.

### Create the add-in project

1. On the Visual Studio menu bar, choose  **File** > **New** > **Project**.
    
2. In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type. 

3. Name the project **HelloWorld**, and then choose the **OK** button.

4. In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.

5. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

     ![PowerPoint tutorial - Visual Studio Solution Explorer window that shows the 2 projects in the HelloWorld solution](../images/powerpoint-tutorial-solution-explorer.png)

### Explore the Visual Studio solution

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### Update code 

Edit the add-in code as follows, to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.

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

2. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

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
