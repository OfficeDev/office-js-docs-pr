---
title: Convert custom functions to the unified manifest
description: Convert an Excel custom functions add-in from the add-in only manifest to the unified manifest for Microsoft 365.
ms.date: 10/01/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Convert custom functions to the unified manifest

This article shows how to convert an existing Excel custom functions add-in from the XML-formatted add-in only manifest to the JSON-formatted unified manifest for Microsoft 365. It supplements [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md) with the manual steps that are specific to custom functions.

The manifest conversion tools don't currently add the custom functions configuration to the generated unified manifest. You must map the custom functions runtime, namespace, metadata URL, and code URLs from the add-in only manifest to the unified manifest.

> [!IMPORTANT]
> The unified manifest isn't supported on every Office version and platform. Before you convert a production add-in, review [Client and platform support](../develop/unified-manifest-overview.md#client-and-platform-support). You might need to maintain and deploy both manifest versions. See [Manage both a unified manifest and an add-in only manifest version of your Office Add-in](../concepts/duplicate-legacy-metaos-add-ins.md).

## Prerequisites

This walkthrough assumes that your project uses Node.js and npm. The add-in can use either a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime) or a [shared runtime](../testing/runtimes.md#shared-runtime). The runtime-specific steps in this article preserve the existing runtime configuration.

The filenames and URLs in this article are examples. Use the corresponding values from your project.

## 1. Prepare the custom functions metadata

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

## 2. Convert the project

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

> [!IMPORTANT]
> The add-in only manifest will be stored in a backup zip file in the root of the project. To update the unified manifest, it's helpful to reference the old manifest's values. Keep it accessible while following this guide.

## 3. Update the manifest schema version

The conversion tool might generate a manifest that uses an old version of the manifest schema. Update it to the latest version. See [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

At the beginning of the generated unified manifest, update `"$schema"` and `"manifestVersion"` to the latest version. The following is an example.

```json
"$schema": "https://developer.microsoft.com/json-schemas/teams/v1.30/MicrosoftTeams.schema.json#",
"manifestVersion": "1.30"
```

Don't change the separate `"version"` property, such as `"version": "1.0.0"`, unless you're also releasing a new version of the add-in.

## 4. Verify permissions

At the root of the manifest, verify that the resource-specific permissions include `Document.ReadWrite.User`. This is necessary for custom functions. This permission is the unified manifest equivalent of `<Permissions>ReadWriteDocument</Permissions>`.

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

## 5. Configure the custom functions runtime

The runtime configuration depends on whether the existing add-in uses a JavaScript-only runtime or a shared runtime. In both cases, get the namespace, metadata URL, and code URLs from the custom functions `<ExtensionPoint>` in the add-in only manifest.

# [JavaScript-only runtime](#tab/javascript-only)

Add a runtime object to the workbook extension's `"runtimes"` array. Configure it as follows.

1. Set `"lifetime"` to `"short"`.
1. Set `"type"` to `"general"`.
1. Set `"code.page"` to the URL referenced by the XML `<Page>` element.
1. Set `"code.script"` to the URL referenced by the XML `<Script>` element.
1. Add a `"customFunctions"` object.
    1. Set both `"namespace.id"` and `"namespace.name"` to the value of the XML `<Namespace>` element. The `id` must remain stable. The `name` is the value shown to users and can be localized.
    1. Set `"metadataUrl"` to the URL referenced by the XML `<Metadata>` element.

The following example shows a JavaScript-only custom functions runtime.

```json
{
  "runtimes": [
    {
      "id": "FunctionsRuntime",
      "type": "general",
      "code": {
        "page": "https://localhost:3000/functions.html",
        "script": "https://localhost:3000/functions.js"
      },
      "lifetime": "short",
      "customFunctions": {
        "namespace": {
          "id": "CONTOSO",
          "name": "CONTOSO"
        },
        "metadataUrl": "https://localhost:3000/functions.json"
      }
    }
  ]
}
```

# [Shared runtime](#tab/shared)

In the workbook extension object, find the runtime generated from the XML `<Runtime>` element. This is normally the runtime with `"lifetime": "long"`.

Configure the runtime as follows.

1. Set `"code.script"` to the URL referenced by the XML `<Script>` element.
1. Add a `"customFunctions"` object to the long-lived runtime. Don't add it to the short-lived runtime used by a ribbon command.
    1. Set both `"namespace.id"` and `"namespace.name"` to the value of the XML `<Namespace>` element. The `id` must remain stable. The `name` is the value shown to users and can be localized.
    1. Set `"metadataUrl"` to the URL referenced by the XML `<Metadata>` element.

The following partial example shows the properties to add or update in the shared runtime.

```json
{
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
}
```

---

> [!NOTE]
> Don't add any custom function ID to the runtime `"actions"` array. The `"actions"` array registers add-in commands. Custom functions are registered by the metadata file and calls to `CustomFunctions.associate`.

## 6. Validate the unified manifest

From the root of a project created with the Yeoman generator or Agents Toolkit, run the following command.

```command&nbsp;line
npm run validate
```

If the project doesn't have a validation script, run the following command. Replace the filename with the path to the unified manifest.

```command&nbsp;line
npx office-addin-manifest validate -p manifest.json
```

Resolve all schema and configuration errors before sideloading. For more validation options, see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md). The manifest reference is at [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

## 7. Sideload and test the converted add-in

Follow [Sideload Office Add-ins that use the unified manifest for Microsoft 365](../testing/sideload-add-in-with-unified-manifest.md) for your project type.

Verify the following behavior.

1. Open a new workbook and confirm that the add-in loads. It may take as long as 2 minutes.
1. Enter a formula that uses each custom function category in your add-in.
1. Confirm that the existing namespace and function names appear in formula autocomplete.
1. Open representative existing workbooks and confirm that their formulas calculate without changes.
1. Test streaming, volatile, cancelable, and dynamic array functions, if the add-in defines them.
1. If the add-in has a task pane, open and close it and confirm that functions continue to calculate.
1. If the add-in shares data between the task pane and custom functions, test that data sharing.
1. Test authentication and external web requests.
1. Test localized function names and descriptions, if the add-in supports localization.

If updated functions don't appear, [clear the Office cache](../testing/clear-cache.md) and sideload the add-in again.

## 8. Plan production deployment

The conversion creates an add-in with a new manifest identity. Don't remove the add-in only manifest version until all targeted clients can install and run the unified manifest version.

Use [Manage both a unified manifest and an add-in only manifest version of your Office Add-in](../concepts/duplicate-legacy-metaos-add-ins.md) to link the versions and hide duplicate UI where supported. Before deployment, test the installation and update experience with representative users, clients, and existing workbooks.

## Troubleshoot the conversion

| Symptom | Check |
| --- | --- |
| Cells show `#NAME?` | Verify the namespace, metadata URL, metadata IDs, and `CustomFunctions.associate` calls. Note that the sideloading process may take several minutes to complete the registration of custom functions. |
| Cells remain `#BUSY!` | Verify `Document.ReadWrite.User`, promise completion, and network requests. Note that the sideloading process may take several minutes to complete the registration of custom functions. |
| Manifest validation rejects `customFunctions` | Confirm that `"customFunctions"` is inside the applicable object in `"runtimes"`, not directly in the extension object. |
| Functions work but ribbon commands don't | Confirm that every ribbon `actionId` matches an `id` in the runtime `"actions"` array. |
| Changes to functions don't appear | [Clear the Office cache](../testing/clear-cache.md) and confirm that the current metadata and script files are served at the manifest URLs. |
| The add-in works on one client but not another | Check unified manifest platform support. For an add-in that uses a shared runtime, also check the SharedRuntime 1.1 requirement set. |

For more help, see [Troubleshoot custom functions](custom-functions-troubleshooting.md).

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Manually create JSON metadata for custom functions](custom-functions-json.md)
- [Office Add-ins with the unified manifest for Microsoft 365](../develop/unified-manifest-overview.md)
