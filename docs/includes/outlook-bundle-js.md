> [!TIP]
>
> - There's no direct method to determine the Outlook profile GUID and mail account encoding used in the **bundle.js** file path. The most effective approach to locate your add-in's **bundle.js** file is to manually inspect each folder until you locate the **Javascript** folder that contains your add-in's ID.
>
> - If the **bundle.js** file doesn't appear in the **Wef** folder and your add-in is installed or sideloaded, restart Outlook. Alternatively, [remove your add-in](../outlook/sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) from Outlook, then [sideload](../outlook/sideload-outlook-add-ins-for-testing.md) it again.
