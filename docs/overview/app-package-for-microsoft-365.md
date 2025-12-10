---
title: App package for Microsoft 365
description:  Learn how add-ins use the app package for Microsoft 365, for packaging, publishing, and management.
ms.date: 12/10/2025
ms.topic: concept-article
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# App package for Microsoft 365

The app package of an app for Microsoft 365 is a zip file that contains a manifest file, two app icons, and possibly additional configuration or localization files. Your app logic and data storage are hosted elsewhere and accessed by the Microsoft 365 host application via HTTPS. You'll submit the app package to your admin to publish to your organization or to Partner Center to publish to Microsoft Marketplace.

At minimum, an app package contains:

- The **app manifest** (`manifest.json`), which describes app configuration, capabilities, required resources, and important attributes.
- A **large full-color icon** (`color.png`), 192 x 192 pixels, to display your agent in the Microsoft 365 Copilot UI and store.
- A **small outline icon** (`outline.png`), 32 x 32 pixels, with a transparent background (not currently used in Copilot, but required to pass validation).

The app package can also contain declarative agent and API plugin definitions, as well as localization files for other supported languages.

:::image type="content" source="../images/app-package.png" alt-text="Diagram showing the anatomy of a Microsoft 365 app package: app manifest (.json file) + icons (color and outline .png files) wrapped in a .zip file." border="false":::

## App manifest

The unified app manifest for Microsoft 365 is a JSON file that describes the functionality and characteristics of your add-in, such as:

- The add-in's display name, description, ID, version, and default locale.

- How the add-in integrates with Office.  

- How the add-in integrates with Copilot (preview).

- The permission level and data access requirements for the add-in.

For a detailed overview of the manifest, see [Office Add-ins with the unified app manifest for Microsoft 365](../develop/unified-manifest-overview.md). For reference documentation, see [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

## App icons

Your app package must include both a color and outline version of your app icon, as .png files. These icons have specific size requirements in order to pass store validation.

For detailed design guidance for color and outline icons for the Microsoft 365 app package, see [Design icons for add-in acquisition and management](../design/microsoft-365-extension-management-icons.md).

### Color icon

The color icon represents your app for Microsoft 365 within the Copilot UI and in-product (Teams, Office, Outlook, Microsoft 365) app stores.

:::row:::
:::column:::

:::image type="content" source="../images/color-icon.png" alt-text="Sample image of an app color icon, showing 192x192 pixels as total icon size with background included, and a central 120x120 pixel space showing the 'Safe region' for the app symbol.":::

:::column-end:::
:::column span="2":::

Your color icon:

- Can be any color.
- Must be 192 x 192 pixels in size.
- Should contain a symbol within 120 x 120 pixels (to allow 36 pixels of padding for [host scenarios where it's cropped](/microsoftteams/platform/concepts/design/design-teams-app-icon-store-appbar#color-icon-architecture)).
- Must sit atop a fully solid or fully transparent square background.

:::column-end:::
:::row-end:::

### Outline icon

The outline icon is used to represent pinned and/or active apps on the Teams app bar. It's not currently used for agents, but still required in order for the app package to pass validation.

:::row:::
:::column:::

:::image type="content" source="../images/outline-icon.png" alt-text="Sample image of an app outline icon, showing 32x32 pixel size and white icon outline with transparent background":::

:::column-end:::
:::column span="2":::

Your outline icon:

- Must be 32 x 32 pixels.
- Must be either white with a transparent background, or transparent with a white background.
- Must not contain additional padding around the symbol.

:::column-end:::
:::row-end:::

## Other configuration and localization files

In addition to the manifest and the two icon files, the app package may also contain some of the following files.

- Localization files that are referenced in the `"localizationInfo"` property of the manifest.
- Copilot declarative agent files that are referenced in the `"copilotAgents"` property.
- Any second-level supplementary files. For example, Copilot declarative agent configuration files sometimes reference second-level supplementary files, such as Copilot plugin configuration files.

## Manually create the app package file

This app package file is usually created for you by the tools you use to create and test your app for Microsoft 365, but there are scenarios in which you create it manually. To do so, use any zip utility to create a zip file that contains the following files.

- The unified manifest, which goes in the root of the zip file.
- The two image files referenced in the `"icons"` property of the manifest.
- Any localization files that are referenced in the `"localizationInfo"` property of the manifest.
- Any declarative agent files that are referenced in the `"copilotAgents"` property.
- Any second-level supplementary files. For example, Copilot declarative agent configuration files sometimes reference second-level supplementary files, such as Copilot plugin configuration files. These should be included too.

> [!IMPORTANT]
> *All of these files must have the same relative path in the zip file as specified in the manifest.* For example, if the path of the two image files is **assets/color.png** and **assets/outline.png**, then you must include an **assets** folder with the two files in the zip package. Second-level files, such as plugin configuration files for declarative agents, must have the same relative path in the zip file as they do in the first-level file that references them. For example, if the relative path of a declarative agent file specified in the manifest is **agents/myAgent.json**, then you must include an **agents** folder in the zip package and put the **myAgent.json** file in it. If the declarative agent file, in turn, gives the relative path of **plugins/myPlugin.json** for a plugin configuration file, then you must include a **plugins** subfolder under the **agents** folder and put the **myPlugin.json** file in it.

To maximize compatibility with Microsoft 365 development tools, we recommend that you keep the files that will be included in the package in a folder called **appPackage** in the root of your project, and that you put the package file in a subfolder named **build** in the **appPackage** folder.

The following are examples of the recommended structure. The structure inside the **\build\appPackage.zip** file must mirror the structure of the **appPackage** folder, except for the **build** folder itself.

```console
\appPackage
    \assets
        color.png
        outline.png
    \build
        appPackage.zip
    manifest.json
```

```console
\appPackage
    \agents
        myAgent.json
        \plugins
            myPlugin.json
    \assets
        color.png
        outline.png
    \build
        appPackage.zip
    \languages
        fr-FR.json
        es-MX.json
    manifest.json
```

> [!NOTE]
> The JSON files that are referenced in the manifest's `"extensions.keyboardShortcuts.keyMappingFiles"` property are *not* included in the app package. They are deployed with the add-in's web application files. For more information, see [Support backward compatibility for add-ins with a unified manifest in Microsoft Marketplace](../design/keyboard-shortcuts.md#support-backward-compatibility-for-add-ins-with-a-unified-manifest-in-microsoft-marketplace).

