To debug your add-in’s initialization sequence, configure your environment so that Microsoft WebView2 (Chromium-based) developer tools automatically open when the add-in launches.

1. Close the Office application where you plan to debug the add-in.
1. Set the `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` environment variable to include the value `--auto-open-devtools-for-tabs`.
1. Open the Office application.
1. Run the add-in.
1. The Microsoft Edge (Chromium-based) developer tools should automatically open. Use the tool the same as you would when debugging a task pane, as specified in [Debug a task pane add-in using Microsoft Edge (Chromium-based) developer tools](../testing/debug-add-ins-using-devtools-edge-chromium.md#debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools).

 > [!NOTE]
 > You may see other instances of the Microsoft Edge (Chromium-based) developer tool auto-opening since this environment variable will affect all WebView2 instances in your system.
