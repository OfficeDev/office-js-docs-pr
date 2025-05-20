> [!TIP]
> If you'll be testing your add-in across multiple environments (for example, in development, staging, demo, etc.), we recommend that you maintain a different manifest file for each environment. In each manifest file, you can:
>
> - Specify the URLs that correspond to the environment.
> - Customize metadata values so that end users are able to identify a sideloaded add-in's corresponding environment. For example:
>
>    - In the unified manifest for Microsoft 365, customize the [`"name"`](/microsoft-365/extensibility/schema/root-name) property of the add-in and the `"label"` properties for various UI controls to indicate the environment.
>    - In the add-in only manifest, customize the `DisplayName` element and and labels within the `Resources` element to indicate the environment.
>
> - Customize the custom functions `namespace` to indicate the environment, if your add-in defines custom functions.
>
> By following this guidance, you'll streamline the testing process and avoid issues that would otherwise occur when an add-in is simultaneously sideloaded for multiple environments.