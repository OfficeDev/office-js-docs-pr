In this step of the tutorial, you'll retrieve metadata for the selected slide.

## Get slide metadata

1. In the **Home.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. In the **Home.js** file, replace `TODO7` with the following code to define the **getSlideMetadata** function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.

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

## Test the add-in

1. Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![A screenshot of Visual Studio with the Start button highlighted](../images/powerpoint-tutorial-start.png)

2. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![A screenshot of Visual Studio with the Show Taskpane button highlighted in the Home ribbon](../images/powerpoint-tutorial-show-taskpane-button.png)

3. In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.

    ![A screenshot of PowerPoint add-in with the Get Slide Metadata button highlighted](../images/powerpoint-tutorial-get-slide-metadata.png)

4. In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![A screenshot of Visual Studio with the Stop button highlighted](../images/powerpoint-tutorial-stop.png)
