You'll begin this tutorial by setting up your development project. 

## Prerequisites

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

## Setup

In this tutorial, you'll create your add-in using Visual Studio.

1. On the Visual Studio menu bar, choose  **File** > **New** > **Project**.
    
2. In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type. 

3. Name the project, and then choose **OK**.

4. In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.

5. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## Add boilerplate code 

Populate the **Home.html** and **Home.js** files with some boilerplate code that you'll modify in subsequent steps of this tutorial.

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
    /// <reference path="Scripts/FabricUI/MessageBanner.js" />

    (function () {
        "use strict";

        var messageBanner;

        // The initialize function must be run each time a new page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: assign event handler for insert-image button.
                // TODO3: assign event handler for insert-text button.
                // TODO5: assign event handler for get-slide-metadata button.
                // TODO7: assign event handler for go-to-slide buttons.
            });
        };

        function insertImageFromBing() {

            //Get image from from webservice. 
            //The service should fetch the photo return it as a base 64 embedded string
            $.ajax({
                url: "/api/Photo/", success: function (result) {
                    insertImageFromBase64String(result);
                }, error: function (xhr, status, error) {

                    showNotification("Error", "Oops, something went wrong.");
                }
            });
        }

        // TODO2: Insert image function. 
    
        // TODO4: Insert text function.

        // TODO6: Get slide metadata function.

        // TODO8: Navigate slides functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notificationHeader").text(header);
            $("#notificationBody").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```
