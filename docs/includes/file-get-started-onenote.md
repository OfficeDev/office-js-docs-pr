# Build your first OneNote add-in

This article walks you through building a simple task pane add-in that adds some text to a OneNote page.

The following image shows the add-in that you'll create.

![The OneNote add-in built from this walkthrough](../images/onenote-first-add-in.png)

<a name="setup"></a>
## Step 1: Set up your dev environment and create an add-in project

Follow the instructions to [Create an Office Add-in using any editor](../get-started/create-an-office-add-in-using-any-editor.md) to install the necessary prerequisites and run the Office Yeoman generator to create a new add-in project. The following table lists  the project attributes to select in the Yeoman generator.

| Option | Value |
|:------|:------|
| New subfolder | (accept the default) |
| Add-in name | OneNote Add-in |
| Supported Office application | (select OneNote) |
| Create new add-in | Yes, I want a new add-in |
| Add [TypeScript](https://www.typescriptlang.org/) | No |
| Choose framework | Jquery |

<a name="develop"></a>
## Step 2: Modify the add-in

You can edit the add-in files using any text editor or IDE. If you haven't tried Visual Studio Code yet, you can [download it for free](https://code.visualstudio.com/) on Linux, Mac OSX, and Windows.

1. Open **index.html** in the project directory. 

2. Replace the `<main>` element with the following code. This adds a text area and a button using [Office UI Fabric components](http://dev.office.com/fabric/components).

      ```html
      <main class="ms-welcome__main">
         <br />
         <p class="ms-font-l">Enter content below</p>
         <div class="ms-TextField ms-TextField--placeholder">
             <textarea id="textBox" rows="5"></textarea>
         </div>
         <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
              <span class="ms-Button-label">Add Outline</span>
              <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
              <span class="ms-Button-description">Adds the content above to the current page.</span>
          </button>
      </main>
      ```

3. Open **app.js** (or app.ts if using TypeScript) in the project directory. Edit the **Office.initialize** function to add a click event to the **Add outline** button, as follows.

      ```js
      // The initialize function is run each time the page is loaded.
      Office.initialize = function (reason) {
         $(document).ready(function () {
             app.initialize();

             // Set up event handler for the UI.
             $('#addOutline').click(addOutlineToPage);
         });
      };
      ```
 
4. Replace the **run** method with the following **addOutlineToPage** method. This gets the content from the text area and adds it to the page.

      ```js
      // Add the contents of the text area to the page.
      function addOutlineToPage() {        
         OneNote.run(function (context) {
            var html = '<p>' + $('#textBox').val() + '</p>';

             // Get the current page.
             var page = context.application.getActivePage();

             // Queue a command to load the page with the title property.             
             page.load('title'); 

             // Add an outline with the specified HTML to the page.
             var outline = page.addOutline(40, 90, html);

             // Run the queued commands, and return a promise to indicate task completion.
             return context.sync()
                 .then(function() {
                     console.log('Added outline to page ' + page.title);
                 })
                 .catch(function(error) {
                     app.showNotification("Error: " + error); 
                     console.log("Error: " + error); 
                     if (error instanceof OfficeExtension.Error) { 
                         console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                     } 
                 }); 
             });
      }
      ```

<a name="test"></a>
## Step 3: Test the add-in on OneNote Online

1. Start the HTTPS server.  

   a. Open a **cmd** prompt / Terminal and go to the add-in project folder. 

   b. Run the following command.

      ```bash
      C:\your-local-path\onenote add-in\> npm start
      ```
2. Install the self-signed certificate as a trusted certificate. You only need to do this one time on your computer for all add-in projects created with the Office Yeoman generator. For more information, see [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

3. Go to [OneNote Online](https://www.onenote.com/notebooks) and open a notebook.

4. Select **Insert > Office Add-ins**. This opens the Office Add-ins dialog.

   - If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then select **Upload My Add-in**.

   - If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**. 
  
   The following image shows the **MY ADD-INS** tab for consumer notebooks.

   <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. In the Upload Add-in dialog, browse to **onenote-add-in-manifest.xml** in your project folder, and then select **Upload**. While testing, your manifest file is stored in the browser's local storage.

6. The add-in opens in an iFrame next to the OneNote page. Enter some text in the text area, and then select **Add outline**. The text you entered is added to the page. 

## Troubleshooting and tips
- You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.

- When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.

   ![Unloaded OneNote object in the debugger](../images/onenote-debug.png)

- You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.

- Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.). 

## See also

- [OneNote JavaScript API programming overview](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API reference](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview](../overview/office-add-ins.md)