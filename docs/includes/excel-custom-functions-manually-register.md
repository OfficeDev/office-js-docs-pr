If the `CONTOSO` namespace isn't available in the autocomplete menu, take the following steps to register the add-in in Excel.

# [Excel on Windows or Mac](#tab/excel-windows)

1. In Excel, select the **Home** > **Add-ins**. Under **My Add-ins**, select **See all**.

1. Under the **MY ADD-INS** section, select **My custom functions add-in** to register it.

# [Excel on the web](#tab/excel-online)

1. In Excel, select **Home** > **Add-ins**, then select **Get Add-ins**.

1. Under the **MY ADD-INS**, select **Manage My Add-ins** and choose **Upload My Add-in**.

1. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.

1. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

1. Try out the new function. In cell **B1**, type the text **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** and press Enter. You should see that the result in cell **B1** is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).

---
