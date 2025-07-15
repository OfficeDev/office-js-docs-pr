> [!TIP]
>
> - For readability, this article refers to the file name as **bundle.js**, but exact name depends on the Office application.
>   - Excel: **bundle_excel.js**
>   - Outlook: **bundle.js**.
>   - PowerPoint: **bundle_powerpoint.js**
>   - Word: **bundle_word.js**
> - There's no direct method to determine the Office profile GUID and account encoding used in the **bundle.js** file path. The most effective approach to locate your add-in's **bundle.js** file is to manually inspect each folder until you locate the **Javascript** folder that contains your add-in's ID.
> - The **bundle.js** file is downloaded to the local **Wef** folder when the add-in is first installed. It's refreshed every time the Office application starts or is restarted. If the **bundle.js** file doesn't appear in the **Wef** folder and your add-in is installed or sideloaded, restart Office. For Outlook, you may need to [remove your add-in](../outlook/sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in), then [sideload](../outlook/sideload-outlook-add-ins-for-testing.md) it again.
