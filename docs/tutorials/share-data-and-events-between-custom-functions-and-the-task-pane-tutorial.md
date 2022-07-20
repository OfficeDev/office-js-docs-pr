---
title: 'Tutorial: Share data and events between Excel custom functions and the task pane'
description: Learn how to share data and events between custom functions and the task pane in Excel.
ms.date: 06/15/2022
ms.prod: excel
ms.localizationpriority: high
---

# Tutorial: Share data and events between Excel custom functions and the task pane

Share global data and send events between the task pane and custom functions of your Excel add-in with a shared runtime. We recommend using a shared runtime for most custom functions scenarios, unless you have a specific reason to use a custom function-only add-in. This tutorial assumes you're familiar with using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create add-in projects. Consider completing the [Excel custom functions tutorial](excel-tutorial-create-custom-functions.md), if you haven't already.

## Create the add-in project

Use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create the Excel add-in project.

- To generate an Excel add-in with custom functions, run the following command.

    ```command&nbsp;line
    yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true
    ```

The generator creates the project and installs supporting Node components.

## Configure the manifest

Follow these steps to configure the add-in project to use a shared runtime.

1. Start Visual Studio Code and open the add-in project you generated.
1. Open the **manifest.xml** file.
1. Replace (or add) the following **\<Requirements\>** section XML to require the [shared runtime requirement set](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets).

    ```xml
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

    After updating, your manifest XML should appear in the following order.

    ```xml
    <Hosts>
      <Host Name="..."/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Find the **\<VersionOverrides\>** section and add the following **\<Runtimes\>** section. The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed. The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the `<bt:Urls>` section near the bottom of the **manifest.xml** file.

    ```xml
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    ```

    > [!IMPORTANT]
    > The **\<Runtimes\>** section must be entered after the `<Host xsi:type="...">` element in the exact order shown in the following XML.

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host xsi:type="...">
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```

    > [!NOTE]
    > If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime) and the conditions for using Microsoft Edge with WebView2 (Chromium-based) are met, it uses that WebView2 control. If the conditions are not met, then it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version. For more information, see [Runtimes](/javascript/api/manifest/runtimes) and [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

1. Find the **\<Page\>** element. Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.

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

1. Open the **webpack.config.js** file.
1. Go to the `plugins:` section.
1. Remove the following `functions.html` plugin if it exists.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Remove the following `commands.html` plugin if it exists.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. If you removed the `functions` or `commands` plugin, add them as `chunks`. The following JavaScript shows the updated entry if you removed both `functions` and `commands` plugins.

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Save your changes and rebuild the project.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > You can also remove the **functions.html** and **commands.html** files. The **taskpane.html** loads the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.

1. Save your changes and run the project. Ensure that it loads and runs with no errors.

   ```command&nbsp;line
   npm run start
   ```

## Share state between custom function and task pane code

Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object. The following instructions show how to share a global variable between custom function and task pane code.

### Create custom functions to get or store shared state

1. In Visual Studio Code open the file **src/functions/functions.js**.
1. On line 1, insert the following code at the very top. This will initialize a global variable named **sharedState**.

    ```js
    window.sharedState = "empty";
    ```

1. Add the following code to create a custom function that stores values to the **sharedState** variable.

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

1. Add the following code to create a custom function that gets the current value of the **sharedState** variable.

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

1. Save the file.

### Create task pane controls to work with global data

1. Open the file **src/taskpane/taskpane.html**.
1. Add the following script element just before the closing `</head>` element.

    ```HTML
    <script src="../functions/functions.js"></script>
    ```

1. After the closing `</main>` element, add the following HTML. The HTML creates two text boxes and buttons used to get or store global data.

    ```HTML
    <ol>
      <li>
        Enter a value to send to the custom function and select
        <strong>Store</strong>.
      </li>
      <li>
        Enter <strong>=CONTOSO.GETVALUE()</strong> into a cell to retrieve it.
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

1. Before the closing `</body>` element, add the following script. This code will handle the button click events when the user wants to store or get global data.

    ```HTML
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

1. Save the file.
1. Build the project.

   ```command line
   npm run build
   ```

### Try sharing data between the custom functions and task pane

Start the project by using the following command.

```command line
npm run start
```

Once Excel starts, you can use the task pane buttons to store or get shared data. Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data. Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.

> [!NOTE]
> Configuring your project as shown in this article will share context between custom functions and the task pane. Calling some Office APIs from custom functions is possible. [See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.

## See also

- [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
