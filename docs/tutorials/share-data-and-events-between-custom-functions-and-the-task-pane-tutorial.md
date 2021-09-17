---
title: "Tutorial: Share data and events between Excel custom functions and the task pane"
description: 'Learn how to share data and events between custom functions and the task pane in Excel.'
ms.date: 09/17/2021
ms.prod: excel
ms.localizationpriority: high
---

# Tutorial: Share data and events between Excel custom functions and the task pane

You can configure your Excel add-in to use a shared runtime. This makes it possible to shared global data, or send events between the task pane and custom functions. For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function. This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects. Consider completing the [Excel custom functions tutorial](excel-tutorial-create-custom-functions.md), if you haven't already.

## Create the add-in project

You'll use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create the Excel add-in project.

- To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.

The generator will create the project and install supporting Node components.

## Configure the manifest

Follow these steps to configure the add-in project to use a shared runtime.

1. Start Visual Studio Code and open the add-in project you generated.
1. Open the **manifest.xml** file.
1. Update the requirements section to use the [shared runtime](../reference/requirement-sets/shared-runtime-requirement-sets.md) instead of the custom function runtime. The XML should appear as follows.

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Find the `<VersionOverrides>` section and add the following `<Runtimes>` section. The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed. The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.

    > [!IMPORTANT]
    > The `<Runtimes>` section must be entered after the `<Host>` element in the exact order shown in the following XML.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
         <Runtimes>
           <Runtime resid="Taskpane.Url" lifetime="long" />
         </Runtimes>
       ...
       </Host>
   ```

1. Find the `<Page>` element. Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Save the **manifest.xml** file.

## Configure the webpack.config.js file

The **webpack.config.js** will build multiple runtime loaders. You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.

1. Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.
1. Open the **webpack.config.js** file.
1. Remove the following **functions.html** plugin code.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Remove the following **commands.html** plugin code.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Add **functions** and **commands** entries to the chunks list as shown next.

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Save your changes and rebuild the project.

   ```command line
   npm run build
   ```

> [!NOTE]
> You can also remove the **functions.html** and **commands.html** files. The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.

1. Save your changes and run the project. Ensure that it loads and runs with no errors.

   ```command line
   npm start
   ```

## Share state between custom function and task pane code

Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object. The following instructions show how to share a global variable between custom function and task pane code.

### Create custom functions to get or store shared state

1. In Visual Studio Code open the file **src/functions/functions.js**.
2. On line 1, insert the following code at the very top. This will initialize a global variable named **sharedState**.

   ```js
   window.sharedState = "empty";
   ```

3. Add the following code to create a custom function that stores values to the **sharedState** variable.

   ```js
   /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
   function storeValue(sharedValue) {
     window.sharedState = sharedValue;
     return "value stored";
   }
   ```

4. Add the following code to create a custom function that gets the current value of the **sharedState** variable.

   ```js
   /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
   function getValue() {
     return window.sharedState;
   }
   ```

5. Save the file.

### Create task pane controls to work with global data

1. Open the file **src/taskpane/taskpane.html**.
2. Add the following script element just before the closing `</head>` element.

   ```html
   <script src="functions.js"></script>
   ```

3. After the closing `</main>` element, add the following HTML. The HTML creates two text boxes and buttons used to get or store global data.

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
     <li>Select <strong>Get</strong> to display the value in the task pane.</li>
   </ol>
   <p>Store new value to shared state</p>
   <div>
     <input type="text" id="storeBox" />
     <button onclick="storeSharedValue()">Store</button>
   </div>

   <p>Get shared state value</p>
   <div>
     <input type="text" id="getBox" />
     <button onclick="getSharedValue()">Get</button>
   </div>
   ```

4. Before the closing `</body>` element, add the following script. This code will handle the button click events when the user wants to store or get global data.

   ```js
   <script>
   function storeSharedValue() {
     let sharedValue = document.getElementById('storeBox').value;
     window.sharedState = sharedValue;
   }

   function getSharedValue() {
     document.getElementById('getBox').value = window.sharedState;
   }
   </script>
   ```

5. Save the file.
6. Build the project

   ```command line
   npm run build
   ```

### Try sharing data between the custom functions and task pane

- Start the project by using the following command.

  ```command line
  npm run start
  ```

Once Excel starts, you can use the task pane buttons to store or get shared data. Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data. Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.

> [!NOTE]
> Configuring your project as shown in this article will share context between custom functions and the task pane. Calling some Office APIs from custom functions is possible. [See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.

## See also

- [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
