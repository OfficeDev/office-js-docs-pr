Add-ins are often cached in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.

You can clear the cache by using the personality menu of any task pane add-in.
- Choose the personality menu. Then choose **Clear Web Cache**.
    > [!NOTE]
    > You must run macOS version 10.13.6 or later to see the personality menu.
    
    ![Screen shot of clear web cache option on personality menu.](../images/mac-clear-cache-menu.png)

You can also clear the cache manually by deleting the contents of the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.

> [!NOTE]
> If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:
>    - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office host (e.g., `Excel`)
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
