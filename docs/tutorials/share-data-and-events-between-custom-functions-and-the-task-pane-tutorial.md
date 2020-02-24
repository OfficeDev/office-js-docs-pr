---
ms.date: 02/20/2020
title: "Tutorial: Share data and events between Excel custom functions and the task pane (preview)"
ms.prod: excel
description: In Excel, share data and events between custom functions and the task pane.
localization_priority: Priority
---

# Tutorial: Share data and events between Excel custom functions and the task pane (preview)

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

You can configure your Excel add-in to use a shared runtime. This will make it possible to shared global data, or send events between the task pane and custom functions.

## Create the add-in project

Use the Yeoman generator to create an Excel add-in project. Run the following command and then answer the prompts with the following answers:

```command line
yo office
```

- Choose a project type: **Excel Custom Functions Add-in project**
- Choose a script type: **JavaScript**
- What do you want to name your add-in? **My Office Add-in**

![Screenshot of answering prompts from yo office to create the add-in project.](../images/yo-office-excel-project.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Configure the manifest

1. Start Visual Studio Code and open the **My Office Add-in** project.
2. Open the **manifest.xml** file.
3. Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section. The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Save your changes and rebuild the project.

   ```command line
   npm run build
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
2. Add the following script element just before the `</head>` element.

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

4. Before the `<body>` element add the following script. This code will handle the button click events when the user wants to store or get global data.

   ```js
   <script>
   function storeSharedValue() {
   let sharedValue = document.getElementById('storeBox').value;
   window.sharedState = sharedValue;
   }

   function getSharedValue() {
   document.getElementById('getBox').value = window.sharedState;
   }</script>
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

Once Excel starts, you can use the task pane buttons to store or get shared data. Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data. Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.

> [!NOTE]
> Configuring your project as shown in this article will share context between custom functions and the task pane. Calling Office APIs from custom functions is not supported in the preview.
