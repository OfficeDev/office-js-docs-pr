If the `CONTOSO` namespace isn't available in the autocomplete menu, take the following steps to register the add-in in Excel.

### [Excel on Windows or Mac](#tab/excel-windows)

1. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.

    :::image type="content" source="../images/select-insert.png" alt-text="Screenshot of the Insert ribbon in Excel on Windows, with the My Add-ins down-arrow highlighted.":::

1. In the list of available add-ins, find the **Developer Add-ins** section and select **My custom functions add-in** to register it.

    :::image type="content" source="../images/excel-cf-tutorial-register.png" alt-text="Screenshot of the Insert ribbon in Excel on Windows, with the Excel Custom Functions add-in highlighted in the My Add-ins list.":::

# [Excel on the web](#tab/excel-online)

1. In Excel, choose the **Insert** tab and then choose **Add-ins**.

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Screenshot of the Insert ribbon in Excel on the web, with the My Add-ins button highlighted.":::

1. Choose **Manage My Add-ins** and select **Upload My Add-in**.

1. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.

1. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

1. Try out the new function. In cell **B1**, type the text **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** and press Enter. You should see that the result in cell **B1** is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).

---
