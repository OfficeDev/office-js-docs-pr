If your project is Node.js-based (that is, not developed with Visual Studio and Internet Information server (IIS)), you can force Office on Windows to use either the EdgeHTML webview control that is provided by Edge Legacy or the Trident webview control that is provided by Internet Explorer to run add-ins, even if you have a combination of Windows and Office versions that would normally use a more recent webview. For more information about which browsers and webviews are used by various combinations of Windows and Office versions, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!NOTE]
> The tool that's used to force the change in webview is supported only in the Beta subscription channel of Microsoft 365. Join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/blog/microsoft365insiderblog/join-the-microsoft-365-insider-program-on-windows/4206638) and select the **Beta Channel** option to access Office Beta builds. See also [About Office: What version of Office am I using?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> Strictly, it's the `webview` switch of this tool (see **Step 2**) that requires the Beta channel. The tool has other switches that don't have this requirement.

1. If your project was *not* created with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) tool, you need to install the office-addin-dev-settings tool. Run the following command in a command prompt.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Specify the webview that you want Office to use with the following command in a command prompt in the root of the project. Replace `<path-to-manifest>` with the relative path, which is just the manifest filename if it's in the root of the project. Replace `<webview>` with either `ie` or `edge-legacy`. Note that the options are named after the browsers in which the webviews originated. The `ie` option means "Trident" and the `edge-legacy` option means "EdgeHTML".

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    The following are examples.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.json edge-legacy
    ```
	
    You should see a message in the command line that the webview type is now set to IE (or Edge Legacy).

1. When you're finished, set Office to resume using the default webview for your combination of Windows and Office versions with the following command.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
