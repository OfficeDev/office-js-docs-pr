---
ms.date: 02/13/2020
title: "Tutorial: Share data and events between Excel custom functions and the task pane (preview)"
ms.prod: excel
description: In Excel, share data and events between custom functions and the task pane.
localization_priority: Priority
---

# Tutorial: Share data and events between Excel custom functions and the task pane (preview)

Excel custom functions and the task pane share global data, and can make function calls into each other. To configure your project so that custom functions can work with the task pane, follow the instructions in this article.

> [!NOTE]
> The features described in this article are currently in preview and subject to change. They are not currently supported for use in production environments. The preview features in this article are only available on Excel on Windows. To try the preview features, you will need to [join Office Insider](https://insider.office.com/join).  A good way to try out preview features is by using an Office 365 subscription. If you don't already have an Office 365 subscription, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Create the add-in project

Use the Yeoman generator to create an Excel add-in project. Run the following command and then answer the prompts with the following answers:

```command&nbsp;line
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
3. Change the `<Requirements>` section to use **CustomFunctionsRuntime** version **1.2** as shown in the following code.
    
    ```xml
    <Requirements>
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section. The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.
    
    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
      <Hosts>
        <Host xsi:type="Workbook">
            <Runtimes>
                <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
            </Runtimes>
            <AllFormFactors>
    ```
    
5. In the `<Page>` element, change the source location from **Functions.Page.Url** to **TaskPaneAndCustomFunction.Url**.

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **TaskPaneAndCustomFunction.Url**.
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. In the `<Action>` section, change the source location from **Taskpane.Url** to **TaskPaneAndCustomFunction.Url**.
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. Add a new **Url id** for **TaskPaneAndCustomFunction.Url** that points to **taskpane.html**.
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. Save your changes and rebuild the project.
    
    ```command&nbsp;line
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
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
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
    
    ```command&nbsp;line
    npm run build 
    ```

### Try sharing data between the custom functions and task pane

- Start the project by using the following command.

    ```command&nbsp;line
    npm run start
    ```

Once Excel starts, you can use the task pane buttons to store or get shared data. Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data. Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.

> [!NOTE]
> Configuring your project as shown in this article will share context between custom functions and the task pane. Calling Office APIs from custom functions is not supported in the preview.

