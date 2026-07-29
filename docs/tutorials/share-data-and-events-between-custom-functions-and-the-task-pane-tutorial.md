---
title: 'Tutorial: Share data and events between Excel custom functions and the task pane'
description: Use a shared runtime to share global data between Excel custom functions and an add-in task pane.
ms.date: 07/28/2026
ms.service: excel
ms.topic: tutorial
ms.localizationpriority: high
ai-usage: ai-assisted
---

# Tutorial: Share data and events between Excel custom functions and the task pane

Use a shared runtime to share global data between your Excel add-in's custom functions and task pane. In this tutorial, you add functions and task pane controls that read and update the same variable.

## Prerequisites

Complete the [Excel custom functions tutorial](excel-tutorial-create-custom-functions.md) and use the add-in that you created. The project must use the **Excel Custom Functions using a Shared Runtime** project type and **JavaScript** script type.

## Share state between custom function and task pane code

### Create custom functions to get or store shared state

1. In Visual Studio Code, open `src/functions/functions.js`.
1. At the beginning of the file, add the following code. This code initializes a global variable named `sharedState`.

    ```js
    window.sharedState = "empty";
    ```

1. Add the following code to create a custom function that stores values to `sharedState`.

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

1. Add the following code to create a custom function that gets the current value of `sharedState`.

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

1. Open `src/taskpane/taskpane.html`.
1. After the closing `</main>` element, add the following HTML to create controls that store and retrieve global data.

    ```html
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

1. Before the closing `</body>` element, add the following script to handle the **Store** and **Get** button events.

    ```html
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

  ```command&nbsp;line
  npm run build
  ```

### Try sharing data between the custom functions and task pane

1. Start the project.

    ```command&nbsp;line
    npm run start
    ```

1. In the task pane, enter a value, and then select **Store**.
1. In an Excel cell, enter `=CONTOSO.GETVALUE()` to retrieve the same value.
1. In another cell, enter `=CONTOSO.STOREVALUE("new value")` to update the shared value.
1. In the task pane, select **Get** to display the updated value.

> [!NOTE]
> A shared runtime also enables custom functions to call some Office APIs. For more information, see [Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md).

When you're ready to stop the development server and uninstall the add-in, run the following command.

```command&nbsp;line
npm run stop
```

## See also

- [Excel custom functions tutorial](excel-tutorial-create-custom-functions.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
