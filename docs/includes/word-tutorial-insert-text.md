In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph in the document.

> [!NOTE]
> This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.

## Code the add-in

1. Open the project in your code editor.
2. Open the file index.html.
3. Replace the `TODO1` with the following markup:

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. Open the app.js file.
5. Replace the `TODO1` with the following code. This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still use the parts of the add-in that are supported by their version of Word.

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. Replace the `TODO2` with the following code:

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. Replace the `TODO3` with the following code. Note the following:
   - Your Word.js business logic will be added to the function that is passed to `Word.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.
   - The `context.sync` method sends all queued commands to Word for execution.
   - The `Word.run` is followed by a `catch` block. This is a best practice that you should always follow. 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. Replace `TODO4` with the following code. Note:
   - The first parameter to the `insertParagraph` method is the text for the new paragraph.
   - The second parameter is the location within the body where the paragraph will be inserted. Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");
    ```

## Test the add-in

1. Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.
2. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.
3. Run the command `npm start` to start a web server running on localhost.
4. Sideload the add-in by using one of the following methods:
    - Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. On the **Home** menu of Word, select **Show Taskpane**.
6. In the task pane, choose **Insert Paragraph**.
7. Make a change in the paragraph.
8. Choose **Insert Paragraph** again. Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the "start" of the document's body.

    ![Word tutorial - Insert Paragraph](../images/word-tutorial-insert-paragraph.png)
