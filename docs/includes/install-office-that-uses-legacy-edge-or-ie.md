Use the following procedure to install either a version of subscription Office that uses the Microsoft Edge Legacy webview (EdgeHTML) to run add-ins or a version that uses Internet Explorer (Trident).

1. Download and install the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).
2. In the folder where you installed the tool (where the `setup.exe` file is located), create a text file with the name `config.xml` and add the following contents.

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

3. Change the `Version` value to `16.0.11929.20946` to install a version that uses Edge Legacy, or to `16.0.10730.20348` to install a version that uses Internet Explorer.
4. Optionally, change the value of `OfficeClientEdition` to `"32"` to install 32-bit Office, and change the `Language ID` value as needed to install Office in a different language.
5. Open a command prompt *as an administrator*.
6. Navigate to the folder with the `setup.exe` and `config.xml` files.
7. Run the following command.

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    This command installs Office. The process may take several minutes.

8. [Clear the Office cache](../testing/clear-cache.md).

> [!IMPORTANT]
> After installation, be sure that you turn off automatic updating of Office, so that Office isn't updated to a version that doesn't use webview you want to work with before you've completed using it. **This can happen within minutes of installation.** Follow these steps.
>
> 1. Start any Office application and open a new document.
> 1. Open the **File** tab on the ribbon, and then select **Office Account** or **Account**.
> 1. In the **Product Information** column, select **Update Options**, and then select **Disable Updates**. If that option is not available, then Office is already configured to not update.
