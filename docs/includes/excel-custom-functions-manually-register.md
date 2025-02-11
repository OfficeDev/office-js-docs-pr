If the `CONTOSO` namespace isn't available in the autocomplete menu, take the following steps to register the add-in in Excel.

# [Excel on the web](#tab/excel-online)

1. Select **Home** > **Add-ins**, then select **More Settings**.

1. On the **Office Add-ins** dialog, select **Upload My Add-in**.

1. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.

1. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

1. Try out the new function. In cell **B1**, type the text **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** and press <kbd>Enter</kbd>. You should see that the result in cell **B1** is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).

# [Excel on Windows or on Mac](#tab/excel-windows)

1. In the Excel ribbon, select **Home** > **Add-ins**.

1. Under the **Developer Add-ins** section, select **My custom functions add-in** to register it.

    :::image type="content" source="../images/excel-cf-select-add-in.png" alt-text="The My Add-ins dialog that shows active add-ins, with the My custom function add-in button highlighted.":::

---
