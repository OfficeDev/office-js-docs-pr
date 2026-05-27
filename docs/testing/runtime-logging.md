---
title: Debug your add-in with runtime logging
description: Learn how to use runtime logging to debug your add-in.
ms.date: 05/18/2026
ms.localizationpriority: medium
---

# Debug your add-in with runtime logging

Use runtime logging to debug your add-in's manifest and several installation errors. This feature helps you identify and fix problems with your manifest that XSD schema validation doesn't catch, such as a mismatch between resource IDs. Runtime logging is especially helpful for debugging add-ins that implement add-in commands and Excel custom functions.

> [!NOTE]
> Runtime logging captures **host-level diagnostics**, such as manifest parsing results, add-in loading errors, and initialization conditions. It does **not** capture your JavaScript `console.log()` output. For general JavaScript debugging, use the developer tools for your platform. See [Debug add-ins using developer tools in Microsoft Edge](debug-add-ins-using-devtools-edge-chromium.md).

> [!IMPORTANT]
> Runtime logging affects performance. Turn it on only when you need to debug problems with your add-in manifest.

## Use runtime logging from the command line

The fastest way to use this logging tool is to enable runtime logging from the command line.

> [!IMPORTANT]
> The office-addin-dev-settings tool isn't supported on Mac. For Mac-specific instructions, see the section [Runtime logging on Mac](#runtime-logging-on-mac).

- To enable runtime logging:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```

- To enable runtime logging and write output to a custom file path:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable <path\to\output.txt>
    ```

    Replace `<path\to\output.txt>` with the path where you want the log written, such as `C:\temp\addin_debug.txt`. This argument only sets the **output file location**. It doesn't filter which add-ins are logged. Runtime logging always applies to all add-ins loaded in the Office runtime on that machine.

    > [!NOTE]
    > When you run `--enable` without a filename, Office writes the log to a default location. Specifying a filename changes *where* the log is written, not *what* is logged.

- To disable runtime logging:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- To display whether runtime logging is enabled:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- To display help within the command line for runtime logging:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## Runtime logging on Mac

1. Open **Terminal** and set a runtime logging preference by using the `defaults` command:

    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>` identifies the host for which to enable runtime logging. `<file_name>` is the name of the text file to which the log is written.

    Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding application.

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

The following example enables runtime logging for Word and then opens the log file.

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE]
> You need to restart Office after running the `defaults` command to enable runtime logging.

To turn off runtime logging, use the `defaults delete` command:

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

The following example turns off runtime logging for Word.

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## Use runtime logging to troubleshoot issues with your manifest

To use runtime logging to troubleshoot issues loading an add-in:

1. [Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.

    > [!NOTE]
    > To minimize the number of messages in the log file, sideload only the add-in that you're testing.

1. If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.

    > [!NOTE]
    > An empty or nearly empty log file is expected when your add-in loads without host-level errors. Runtime logging only records manifest and loading diagnostics. It doesn't contain entries if your add-in loads correctly. If you're looking for JavaScript `console.log()` output, use the developer tools for your platform instead.

1. Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.

## Known issues with runtime logging

You might see messages in the log file that are confusing or that are classified incorrectly. For example:

- The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.

- If you see the message `Unexpected Add-in is missing required manifest fields    DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you're debugging.

- Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.

## See also

- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)
- [Clear the Office cache](clear-cache.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug add-ins using developer tools in Microsoft Edge](debug-add-ins-using-devtools-edge-chromium.md)
- [Runtimes in Office Add-ins](runtimes.md)
