If your project is node.js-based (that is, not developed with Visual Studio and Internet Information server (IIS)), you can force Office on Windows to use Edge Legacy or Internet Explorer to run add-ins, even if you have a combination of Windows and Office versions that would normally use a more recent browser. For more information about which browsers are used by various combinations of Windows and Office versions, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

1. If your project was *not* created with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) tool, you need to install the office-addin-dev-settings tool. Run the following command in a command prompt.

   > [!NOTE]
   > This `webview` switch of this tool that you use (see the next step) is supported only in the Beta subscription channel of Microsoft 365. Join the [Office Insider program](https://insider.office.com/join/windows) and select the **Beta Channel** option to access Office beta builds. See also [About Office: What version of Office am I using?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Specify the browser that you want Office to use with the following command in a command prompt in the root of the project. Replace `<path-to-manifest>` with the relative path, which is just the manifest filename if it is in the root of the project. Replace `<webview>` with either `ie` or `edge-legacy`.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    The following is an example.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    You should see a message in the command line that the webview type is now set to IE (or Edge Legacy).

1. When you're finished, set Office to resume using the default browser for your combination of Windows and Office versions with the following command.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
