---
title: 'Tutorial: Share data and events between Excel custom functions and the task pane'
description: Learn how to share data and events between custom functions and the task pane in Excel.
ms.date: 03/20/2023
ms.service: excel
ms.localizationpriority: high
---

# Tutorial: Share data and events between Excel custom functions and the task pane

Share global data and send events between the task pane and custom functions of your Excel add-in with a shared runtime.

## Share a state between custom function and task pane code

The following instructions show how to share a global variable between custom function and task pane code. This tutorial assumes that you've completed the [Excel custom functions tutorial](excel-tutorial-create-custom-functions.md), with a **Excel Custom Functions using a Shared Runtime** project using the script type **JavaScript**. Use the add-in you created in that tutorial to complete the following instructions.

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

- Start the project by using the following command.

    ```command&nbsp;line
    npm run start
    ```

Once Excel starts, you can use the task pane buttons to store or get shared data. Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data. Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.

> [!NOTE]
> Calling some Office APIs from custom functions using a shared runtime is possible. [See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.

When you're ready to stop the dev server and uninstall the add-in, run the following command.

```command&nbsp;line
npm run stop
```

## See also

- [Excel custom functions tutorial](excel-tutorial-create-custom-functions.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
