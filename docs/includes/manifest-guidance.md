> [!TIP]
> If you'll be testing your add-in across multiple environments (for example, in development, staging, demo, etc.), we recommend that you create a different manifest file for each environment. In the manifest file for each environment, you can:
> - Specify the URLs that correspond to the environment.
> - Customize metadata values like `DisplayName` and labels within `Resources` to clearly indicate the environment (so that it'll be clear which environment a sideloaded add-in corresponds to). 
> - Specify a custom functions `namespace` that identifies the environment (if your add-in defines custom functions). 
> By following this guidance, you'll avoid issues that would otherwise occur when an add-in is simultaneously sideloaded for multiple environments.