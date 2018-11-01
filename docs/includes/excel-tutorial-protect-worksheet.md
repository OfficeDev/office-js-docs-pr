In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.

> [!NOTE]
> This page describes an individual step of the Excel add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.

## Configure the manifest to add a second ribbon button

1. Open the manifest file **my-office-add-in-manifest.xml**.
2. Find the `<Control>` element. This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in. We're going to add a second button to the same group on the **Home** ribbon. In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file. There's only one other button in the manifest, so this isn't difficult. Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection". When you are done, the entire start Control tag should look like the following:

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. The next three `TODO`s set "resid"s, which is short for resource ID. A resource is a string, and you'll create these three strings in a later step. For now, you need to give IDs to the resources. The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. The `SuperTip` element defines the tool tip for the button. The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel". The tool tip description will be "Click to turn protection of the worksheet on and off". But the `ID` should be "ProtectionButtonToolTip". So, when you are done, the whole `SuperTip` markup should look like the following code: 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`. 

6. The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step. So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions. The start `Action` tag should look like the following code:
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    The entire `Control` markup should now look like the following:

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. Scroll down to the `Resources` section of the manifest.

9. Add the following markup as a child of the `bt:ShortStrings` element.

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. Add the following markup as a child of the `bt:LongStrings` element.

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. Be sure to save the file.

## Create the function that protects the sheet

1. Open the file \function-file\function-file.js.

2. The file already has an Immediately Invoked Function Expression (IFFE). No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body. (But do not delete it. The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code. Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`. This is a requirement for all add-in commands of type **ExecuteFunction**. It signals the Office host application that the function has finished and the UI can become responsive again.

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
		args.completed();
    }
    ```

3. Replace `TODO1` with the following code. This code uses the worksheet object's protection property in a standard toggle pattern. The `TODO2` will be explained in the next section.

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## Add code to fetch document properties into the task pane's script objects

In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value. It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`. Only then can `sheet.protection.protected` be called without causing an exception to be thrown. This fetching process has three steps:

   1. Queue a command to load (that is; fetch) the properties that your code needs to read.
   2. Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.
   3. Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.

These steps must be completed whenever your code needs to *read* information from the Office document.

1. In the `toggleProtection` function, replace `TODO2` with the following code. Note:
   - Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.
   - To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed. 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.
3. Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.
4. Replace `TODO4` with the following code. Note:
   - Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.
   - The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.

    ```javascript
    .then(context.sync);
    ```

   When you are done, the entire function should look like the following:

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```


## Configure the script-loading HTML file

Open the /function-file/function-file.html file. This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button. Its purpose is to load the JavaScript method that should run when the button is pushed. You are not going to change this file. Simply note that the second `<script>` tag loads the functionfile.js.

   > [!NOTE]
   > The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane. If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the function-file.js at all. 

## Test the add-in

1. Close all Office applications, including Excel. 
2. Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host. 
    - For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
    - For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.
3. If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`. You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.
4. Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods. *You should overwrite the previous copy of the manifest file.*
    - Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
7. Open any worksheet in Excel.
8. On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below. 
9. Choose a cell as you would if you wanted to change its content. You get an error telling you that the worksheet is protected.
10. Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.

    ![Excel tutorial - Ribbon with Protection On](../images/excel-tutorial-ribbon-with-protection-on.png)
