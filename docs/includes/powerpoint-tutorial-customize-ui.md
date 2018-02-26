In this step of the tutorial, you'll customize the task pane user interface (UI).

## Customize the task pane UI 

1. In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane.

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint Add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.

## Test the add-in

1. Using Visual Studio, test the PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.

    ![A screenshot of Visual Studio with the Start button highlighted](../images/powerpoint-tutorial-start.png)

2. In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![A screenshot of Visual Studio with the Show Taskpane button highlighted in the Home ribbon](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Notice that the task pane now contains a header section and title, and no longer contains a footer section.

    ![A screenshot of PowerPoint add-in with the Insert Image button highlighted](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![A screenshot of Visual Studio with the Stop button highlighted](../images/powerpoint-tutorial-stop.png)

