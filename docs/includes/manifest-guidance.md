> [!TIP]
> If you'll be testing your add-in across multiple environments (for example, in development, staging, demo, etc.), we recommend that you maintain a different manifest file for each environment. In each manifest file, you can:
> - Specify the URLs that correspond to the environment.
> - Customize metadata values like `DisplayName` and labels within `Resources` to indicate the environment, so that end users will be able to identify a sideloaded add-in's corresponding environment. 
> - Customize the custom functions `namespace` to indicate the environment, if your add-in defines custom functions.
> 
> By following this guidance, you'll streamline the testing process and avoid issues that would otherwise occur when an add-in is simultaneously sideloaded for multiple environments.