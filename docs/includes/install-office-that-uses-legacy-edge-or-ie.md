Use the following procedure to install either a version of Office (downloaded from a Microsoft 365 subscription) that uses the Microsoft Edge Legacy webview (EdgeHTML) to run add-ins or a version that uses Internet Explorer (Trident).

1. In any Office application, open the **File** tab on the ribbon, and then select **Office Account** or **Account**. Select the **About _host-name_** button (for example, **About Word**).
1. On the dialog that opens, find the full xx.x.xxxxx.xxxxx build number and make a copy of it somewhere.
1. Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).
1. Run the downloaded file to extract the tool. You are prompted to choose where to install the tool.
1. In the folder where you installed the tool (where the `setup.exe` file is located), create a text file with the name `config.xml` and add the following contents.

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. Change the `Version` value.

    - To install a version that uses EdgeHTML, change it to `16.0.11929.20946`.
    - To install a version that uses Trident, change it to `16.0.10730.20348`.

1. Optionally, change the value of `OfficeClientEdition` to `"32"` to install 32-bit Office, and change the `Language ID` value as needed to install Office in a different language.
1. Open a command prompt *as an administrator*.
1. Navigate to the folder with the `setup.exe` and `config.xml` files.
1. Run the following command.

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    This command installs Office. The process may take several minutes.

1. [Clear the Office cache](../testing/clear-cache.md).

> [!IMPORTANT]
> After installation, be sure that you turn off automatic updating of Office, so that Office isn't updated to a version that doesn't use webview you want to work with before you've completed using it. **This can happen within minutes of installation.** Follow these steps.
>
> 1. Start any Office application and open a new document.
> 1. Open the **File** tab on the ribbon, and then select **Office Account** or **Account**.
> 1. In the **Product Information** column, select **Update Options**, and then select **Disable Updates**. If that option isn't available, then Office is already configured to not update automatically.

When you are finished using the old version of Office, reinstall your newer version by editing the `config.xml` file and changing the `Version` to the build number that you copied earlier. Then repeat the `setup.exe /configure config.xml` command in an administrator command prompt. Optionally, re-enable automatic updates.
