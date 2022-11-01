Add-ins are often cached in Office on Mac for performance reasons. Normally, the cache is cleared by reloading the add-in. If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.

### Use the personality menu to clear the cache

You can clear the cache by using the personality menu of any task pane add-in. However, because the personality menu isn't supported in Outlook add-ins, you can try the option to [clear the cache manually](#clear-the-cache-manually) if you're using Outlook.

- Choose the personality menu. Then choose **Clear Web Cache**.
    > [!NOTE]
    > You must run macOS Version 10.13.6 or later to see the personality menu.

    ![Screenshot of clear web cache option on personality menu.](../images/mac-clear-cache-menu.png)

### Clear the cache manually

You can also clear the cache manually by deleting the contents of the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder. Look for this folder via terminal.

> [!NOTE]
> If that folder doesn't exist, check for the following folders via terminal and if found, delete the contents of the folder.
>
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office application (e.g., `Excel`)
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office application (e.g., `Excel`)
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
>
> To look for these folders via Finder, you must set Finder to show hidden files. Finder displays the folders inside the **Containers** directory by product name, such as **Microsoft Excel** instead of **com.microsoft.Excel**.