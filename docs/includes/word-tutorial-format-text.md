In this step of the tutorial, you'll use change the font of text, and use both built-in and custom styles on the text.

> [!NOTE]
> This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.

## Apply a built-in style to text

1. Open the project in your code editor. 
2. Open the file index.html.
3. Just below the `div` that contains the `insert-paragraph` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. Open the app.js file.

5. Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. Just below the `insertParagraph` function, add the following function:

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

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

7. Replace `TODO1` with the following code. Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## Apply a custom style to text

1. Open the file index.html.
2. Below the `div` that contains the `apply-style` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. Open the app.js file.

4. Below the line that assigns a click handler to the `apply-style` button, add the following code:

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. Below the `applyStyle` function add the following function.

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

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

7. Replace `TODO1` with the following code. Note that the code applies a custom style that does not exist yet. You will create a style with the name **MyCustomStyle** when you test the add-in below.

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## Change the font of text

1. Open the file index.html.
2. Below the `div` that contains the `apply-custom-style` button, add the following markup:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. Open the app.js file.

4. Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:

    ```js
    $('#change-font').click(changeFont);
    ```

5. Below the `applyCustomStyle` function add the following function.

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

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

7. Replace `TODO1` with the following code. Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## Test the add-in

1. If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.

     > [!NOTE]
     > Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.

2. Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.
3. Run the command `npm start` to start a web server running on localhost.   
4. Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.
5. Be sure there are at least three paragraphs in the document. You can choose **Insert Paragraph** three times. *Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*
6. In Word, create a custom style named "MyCustomStyle". It can have any formatting that you want.
7. Choose the **Apply Style** button. The first paragraph will be styled with the built-in style **Intense Reference**.
8. Choose the **Apply Custom Style** button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)
9. Choose the **Change Font** button. The font of the second paragraph changes to 18 pt., bold, Courier New.

    ![Word tutorial - Apply Styles and Font](../images/word-tutorial-apply-styles-and-font.png)
