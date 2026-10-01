---
title: Convert custom functions to the unified manifest
description: Convert an Excel custom functions add-in from the add-in only manifest to the unified manifest for Microsoft 365.
ms.date: 09/30/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Convert custom functions to the unified manifest

This article shows how to convert an existing Excel custom functions add-in from the XML-formatted add-in only manifest to the JSON-formatted unified manifest for Microsoft 365. It supplements [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md) with the manual steps that are specific to custom functions.

The manifest conversion tools don't currently add the custom functions configuration to the generated unified manifest. You must map the custom functions runtime, namespace, and metadata URL from the add-in only manifest to the unified manifest.

> [!IMPORTANT]
> The unified manifest isn't supported on every Office version and platform. Before you convert a production add-in, review [Client and platform support](../develop/unified-manifest-overview.md#client-and-platform-support). You might need to maintain and deploy both manifest versions.

## Prerequisites

This walkthrough assumes that your add-in meets the following conditions.

- The add-in uses a [shared runtime](../testing/runtimes.md#shared-runtime), which is the recommended runtime for custom functions.
- The project uses Node.js and npm.
- The add-in has a valid add-in only manifest and can be sideloaded successfully.
- The custom functions metadata is stored in a JSON file, such as **functions.json**.

If your add-in uses the JavaScript-only runtime, you can use the preparation and conversion steps in this article. After conversion, [configure the unified manifest to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) before you configure the custom functions runtime in step 6.

The filenames and URLs in this article are examples. Use the corresponding values from your project.

## 1. Record the custom functions configuration

Before running a conversion tool, record the values that configure custom functions in the add-in only manifest. The following example shows the relevant parts of a typical shared-runtime manifest.

```xml
<Requirements>
  <Sets DefaultMinVersion="1.1">
    <Set Name="SharedRuntime" MinVersion="1.1"/>
  </Sets>
</Requirements>

<Permissions>ReadWriteDocument</Permissions>

<VersionOverrides ...>
  <Hosts>
    <Host xsi:type="Workbook">
      <Runtimes>
        <Runtime resid="Shared.Url" lifetime="long"/>
      </Runtimes>
      <AllFormFactors>
        <ExtensionPoint xsi:type="CustomFunctions">
          <Script>
            <SourceLocation resid="Functions.Script.Url"/>
          </Script>
          <Page>
            <SourceLocation resid="Shared.Url"/>
          </Page>
          <Metadata>
            <SourceLocation resid="Functions.Metadata.Url"/>
          </Metadata>
          <Namespace resid="Functions.Namespace"/>
        </ExtensionPoint>
      </AllFormFactors>
    </Host>
  </Hosts>
  <Resources>
    <bt:Urls>
      <bt:Url id="Functions.Script.Url"
              DefaultValue="https://localhost:3000/functions.js"/>
      <bt:Url id="Functions.Metadata.Url"
              DefaultValue="https://localhost:3000/functions.json"/>
      <bt:Url id="Shared.Url"
              DefaultValue="https://localhost:3000/taskpane.html"/>
    </bt:Urls>
    <bt:ShortStrings>
      <bt:String id="Functions.Namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
  </Resources>
</VersionOverrides>
```

Record the resolved values, rather than only the resource IDs.

| Add-in only manifest setting | Example resolved value |
| --- | --- |
| Runtime page, referenced by `<Runtime>` and `<Page>` | `https://localhost:3000/taskpane.html` |
| Function script, referenced by `<Script>` | `https://localhost:3000/functions.js` |
| Metadata file, referenced by `<Metadata>` | `https://localhost:3000/functions.json` |
| Namespace, referenced by `<Namespace>` | `CONTOSO` |
| Runtime lifetime | `long` |
| Permission | `ReadWriteDocument` |

Also record any locale-specific namespace and metadata overrides. This walkthrough configures the default locale. Don't retire the add-in only manifest version until you've validated the localized experience in the unified manifest version.

## 2. Prepare the custom functions metadata

The unified manifest enforces some metadata requirements that older Office clients or Microsoft Marketplace submissions might not have enforced for the add-in only manifest.

1. Open the custom functions metadata file.
1. Verify that every function `id` and `name` has at least three characters.
1. Verify that every function object has a `result` property.
1. Keep the existing function IDs and names. Changing them can prevent formulas in existing workbooks from resolving.
1. Verify that every metadata ID is associated with its JavaScript or TypeScript implementation. For example:

   ```javascript
   CustomFunctions.associate("ADD", add);
   ```

For all metadata requirements, see [Custom functions naming and localization](custom-functions-naming.md) and [Manually create JSON metadata for custom functions](custom-functions-json.md).

After making any corrections, validate and sideload the add-in only manifest again. Resolve any problems before continuing.

## 3. Convert the project

The conversion command depends on how the project was created.

# [Yeoman generator project](#tab/yo-office)

From the root of the project, run the following command. Replace the manifest path with the path to your add-in only manifest.

```command&nbsp;line
npx office-addin-project convert -m manifest.xml
```

This command converts the manifest and updates the project's npm configuration. It puts the original project files in a backup zip file.

# [Other Node.js and npm project](#tab/other-project)

From the root of the project, run the following command. Replace the manifest path with the path to your add-in only manifest.

```command&nbsp;line
npx office-addin-manifest-converter convert manifest.xml
```

This command creates the unified manifest in a subfolder named after the add-in only manifest.

---

Complete the general post-conversion steps in [Edit the new unified manifest](../develop/convert-xml-to-json-manifest.md#edit-the-new-unified-manifest), including adding the required developer URLs. Don't sideload the add-in yet.

## 4. Update the manifest schema version

The conversion tool might generate a manifest that uses schema version 1.17. The `customFunctions.metadataUrl` property isn't available in that schema version.

At the beginning of the generated unified manifest, update `"$schema"` and `"manifestVersion"` to version 1.30.

```json
"$schema": "https://developer.microsoft.com/json-schemas/teams/v1.30/MicrosoftTeams.schema.json#",
"manifestVersion": "1.30"
```

Don't change the separate `"version"` property, such as `"version": "1.0.0"`, unless you're also releasing a new version of the add-in.

## 5. Configure the extension requirements and permissions

Open the generated unified manifest and find the object in the `"extensions"` array that has `"workbook"` in its `"requirements.scopes"` array.

Ensure that this extension-level requirements object includes the SharedRuntime 1.1 requirement set. The conversion tool normally creates this configuration from the `<Requirements>` element of the add-in only manifest.

```json
"requirements": {
  "scopes": [
    "workbook"
  ],
  "capabilities": [
    {
      "name": "SharedRuntime",
      "minVersion": "1.1"
    }
  ]
}
```

At the root of the manifest, verify that the resource-specific permissions include `Document.ReadWrite.User`. This permission is the unified manifest equivalent of `<Permissions>ReadWriteDocument</Permissions>`.

```json
"authorization": {
  "permissions": {
    "resourceSpecific": [
      {
        "name": "Document.ReadWrite.User",
        "type": "Delegated"
      }
    ]
  }
}
```

## 6. Configure the custom functions runtime

In the workbook extension object, find the runtime generated from the `<Runtime>` element. For a shared-runtime project created by the Yeoman generator, this is normally the runtime with `"lifetime": "long"`. The conversion tool might name it `"runtime_1"` and create a second, short-lived runtime for a ribbon command.

Add `"code.script"` and a `"customFunctions"` object to the long-lived runtime. Keep the runtime's generated ID and requirements. The following example shows the relevant part of a typical manifest generated from a Yeoman custom functions project.

```json
"runtimes": [
  {
    "requirements": {
      "capabilities": [
        {
          "name": "AddinCommands",
          "minVersion": "1.1"
        }
      ],
      "formFactors": [
        "desktop"
      ]
    },
    "id": "runtime_1",
    "type": "general",
    "code": {
      "page": "https://localhost:3000/taskpane.html",
      "script": "https://localhost:3000/functions.js"
    },
    "lifetime": "long",
    "customFunctions": {
      "namespace": {
        "id": "CONTOSO",
        "name": "CONTOSO"
      },
      "metadataUrl": "https://localhost:3000/functions.json"
    }
  },
  {
    "requirements": {
      "capabilities": [
        {
          "name": "AddinCommands",
          "minVersion": "1.1"
        }
      ],
      "formFactors": [
        "desktop"
      ]
    },
    "id": "runtime_2",
    "type": "general",
    "code": {
      "page": "https://localhost:3000/taskpane.html"
    },
    "lifetime": "short",
    "actions": [
      {
        "id": "ButtonId1_1",
        "type": "openPage",
        "displayName": "ButtonId1_1",
        "view": "ButtonId1"
      }
    ]
  }
]
```

Adapt the example as follows.

- Add the custom functions configuration to the runtime with `"lifetime": "long"`. Don't add it to the short-lived runtime used by the ribbon command.
- Set `"code.page"` to the URL referenced by both the XML `<Runtime>` and `<Page>` elements.
- Set `"code.script"` to the URL referenced by the XML `<Script>` element.
- Keep `"lifetime"` set to `"long"` to preserve the shared runtime.
- Set both `"customFunctions.namespace.id"` and `"customFunctions.namespace.name"` to the existing namespace. The `id` must remain stable. The `name` is the value shown to users and can be localized.
- Set `"customFunctions.metadataUrl"` to the URL referenced by the XML `<Metadata>` element.
- If the runtime has an `"actions"` array for ribbon commands, keep the actions created by the conversion tool. Ensure that each ribbon control's `actionId` matches a runtime action `id`.

The extension-level and runtime-level `"requirements"` objects have different purposes. The extension-level `SharedRuntime` capability controls whether the add-in can be installed. A runtime-level requirements object filters only that runtime. You don't need to move or duplicate the `SharedRuntime` capability in the long-lived runtime when it is already present in the extension-level requirements.

> [!NOTE]
> Don't add every custom function ID to the runtime `"actions"` array. The `"actions"` array registers add-in commands. Custom functions are registered by the metadata file and calls to `CustomFunctions.associate`.

The following table summarizes the custom functions mapping.

| Add-in only manifest | Unified manifest |
| --- | --- |
| `<Set Name="SharedRuntime" MinVersion="1.1"/>` | `requirements.capabilities` entry for `SharedRuntime` 1.1 |
| `<Runtime resid="..." lifetime="long">` | Runtime `code.page` and `lifetime` |
| Custom Functions `<Script>` | Runtime `code.script` |
| Custom Functions `<Page>` | Runtime `code.page` |
| Custom Functions `<Metadata>` | Runtime `customFunctions.metadataUrl` |
| Custom Functions `<Namespace>` | Runtime `customFunctions.namespace` |
| `<Permissions>ReadWriteDocument</Permissions>` | `Document.ReadWrite.User` resource-specific permission |

## 7. Preserve the metadata file

Keep the custom functions metadata file in the converted project and continue serving it from the URL in `"metadataUrl"`. Don't copy the complete metadata file into the manifest.

The `"customFunctions"` object in the manifest identifies the namespace and metadata location. The metadata file continues to define each function's ID, name, parameters, result, descriptions, and other options.

If your build generates the metadata file from JSDoc comments, run the normal build and verify that the generated file is available at the configured URL. For more information, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

## 8. Validate the unified manifest

From the root of a project created with the Yeoman generator or Agents Toolkit, run the following command.

```command&nbsp;line
npm run validate
```

If the project doesn't have a validation script, run the following command. Replace the filename with the path to the unified manifest.

```command&nbsp;line
npx office-addin-manifest validate -p manifest.json
```

Resolve all schema and configuration errors before sideloading. For more validation options, see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).

## 9. Sideload and test the converted add-in

Follow [Sideload Office Add-ins that use the unified manifest for Microsoft 365](../testing/sideload-add-in-with-unified-manifest.md) for your project type.

Verify the following behavior.

1. Open a new workbook and confirm that the add-in loads.
1. Enter a formula that uses each custom function category in your add-in.
1. Confirm that the existing namespace and function names appear in formula autocomplete.
1. Open representative existing workbooks and confirm that their formulas calculate without changes.
1. Test streaming, volatile, cancelable, and dynamic array functions, if the add-in defines them.
1. Open and close the task pane and confirm that functions continue to calculate.
1. Test any data shared between the task pane and custom functions.
1. Test authentication and external web requests.
1. Test localized function names and descriptions, if the add-in supports localization.

If updated functions don't appear, [clear the Office cache](../testing/clear-cache.md) and sideload the add-in again.

## 10. Plan production deployment

The conversion creates an add-in with a new manifest identity. Don't remove the add-in only manifest version until all targeted clients can install and run the unified manifest version.

Use [Manage both a unified manifest and an add-in only manifest version of your Office Add-in](../concepts/duplicate-legacy-metaos-add-ins.md) to link the versions and hide duplicate UI where supported. Before deployment, test the installation and update experience with representative users, clients, and existing workbooks.

## Troubleshoot the conversion

| Symptom | Check |
| --- | --- |
| Cells show `#NAME?` | Verify the namespace, metadata URL, metadata IDs, and `CustomFunctions.associate` calls. |
| Cells remain `#BUSY!` | Verify `Document.ReadWrite.User`, promise completion, and network requests. |
| Manifest validation rejects `customFunctions` | Confirm that `"customFunctions"` is inside the applicable object in `"runtimes"`, not directly in the extension object. |
| Functions work but ribbon commands don't | Confirm that every ribbon `actionId` matches an `id` in the runtime `"actions"` array. |
| Changes to functions don't appear | Clear the Office cache and confirm that the current metadata and script files are served at the manifest URLs. |
| The add-in works on one client but not another | Check unified manifest platform support and the SharedRuntime 1.1 requirement set. |

For more help, see [Troubleshoot custom functions](custom-functions-troubleshooting.md).

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Manually create JSON metadata for custom functions](custom-functions-json.md)
- [Office Add-ins with the unified manifest for Microsoft 365](../develop/unified-manifest-overview.md)
