UI-less code, such as add-in commands, or code running while the task pane is not visible, will not be able to attach to a debugger by default. You'll need to use [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell) to run a couple commands to enable debugging.

1. Run the following command to get information for the **Microsoft.Win32WebViewHost** app package.
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    The command lists app package information similar to the following output.
    
    ```powershell
    Name              : Microsoft.Win32WebViewHost
    Publisher         : CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
    Architecture      : Neutral
    ResourceId        : neutral
    Version           : 10.0.18362.449
    PackageFullName   : Microsoft.Win32WebViewHost_10.0.18362.449_neutral_neutral_cw5n1h2txyewy
    InstallLocation   : C:\Windows\SystemApps\Microsoft.Win32WebViewHost_cw5n1h2txyewy
    IsFramework       : False
    PackageFamilyName : Microsoft.Win32WebViewHost_cw5n1h2txyewy
    PublisherId       : cw5n1h2txyewy
    IsResourcePackage : False
    IsBundle          : False
    IsDevelopmentMode : False
    NonRemovable      : True
    IsPartiallyStaged : False
    SignatureKind     : System
    Status            : Ok
    ```
    
2. Run the following command to enabled debugging. Use the value for the **PackageFullName** listed from the previous command.
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. If Office was already running, close and restart Office so that it picks up the debugging change.