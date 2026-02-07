---
title: Configure custom functions with the unified manifest for Microsoft 365
description: Learn how to configure Excel custom functions using the unified manifest for Microsoft 365, including namespace configuration, metadata references, and runtime setup.
ms.date: 02/06/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Configure custom functions with the unified manifest for Microsoft 365

Custom functions enable developers to add new functions to Excel by defining them in JavaScript or TypeScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. This article explains how to configure custom functions using the unified manifest for Microsoft 365.

For an introduction to custom functions, see [Create custom functions in Excel](custom-functions-overview.md).

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

> [!TIP]
> For instructions on configuring custom functions with the add-in only manifest, see [Create custom functions in Excel](custom-functions-overview.md) and [Manually create JSON metadata for custom functions](custom-functions-json.md).

## Prerequisites

- Familiarity with custom functions. See [Custom functions overview](custom-functions-overview.md).
- Familiarity with the unified manifest. See [Office Add-ins with the unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).
- Office version 2304 (Build 16320.00000) or later for unified manifest support.

## Custom functions in the unified manifest

The unified manifest uses the `customFunctions` extension property to configure custom functions. This configuration includes:

1. **Namespace configuration**: Defines the namespace for your custom functions.
2. **Metadata reference**: Points to the JSON metadata file that describes your functions.
3. **Runtime configuration**: Specifies how custom functions execute.

## Configure the customFunctions extension

### Step 1: Define the extension requirements

Open your unified manifest file (manifest.json) and ensure the `extensions` array includes an object with the following requirements:

```json
{
  "extensions": [
    {
      "requirements": {
        "scopes": ["workbook"],
        "capabilities": [
          {
            "name": "CustomFunctionsRuntime",
            "minVersion": "1.1"
          }
        ]
      }
    }
  ]
}
```

**Key properties**:

- `scopes`: Set to `["workbook"]` to specify this extension applies to Excel.
- `capabilities`: Identifies the CustomFunctionsRuntime 1.1 requirement set as the minimum version needed.

### Step 2: Configure the namespace

Add the `customFunctions` object with namespace configuration:

```json
{
  "extensions": [
    {
      "requirements": { /* as above */ },
      "customFunctions": {
        "namespace": {
          "id": "CONTOSO",
          "name": "CONTOSO"
        }
      }
    }
  ]
}
```

**Namespace properties**:

- `id`: Unique identifier used internally (required). This value must remain stable to avoid breaking existing workbooks.
- `name`: Display name shown to users in Excel (required). This value can be localized.

**Naming guidelines**:

- Use uppercase for consistency with Excel's built-in functions.
- Keep names short and memorable (for example, "CONTOSO", "MYCOMPANY").
- Must be unique across all add-ins the user has installed.
- Follow the naming guidelines in [Custom functions naming and localization](custom-functions-naming.md).

### Step 3: Reference the JSON metadata file

Add the `metadataUrl` property to point to your functions metadata file:

```json
{
  "extensions": [
    {
      "requirements": { /* as above */ },
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

**metadataUrl requirements**:

- Must be a complete HTTPS URL (HTTP may work in development environments).
- Maximum length: 2048 characters.
- Points to a JSON file describing function signatures, parameters, and return types.

For details on creating the metadata JSON file, see [Manually create JSON metadata for custom functions](custom-functions-json.md) or [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

### Step 4: Configure the runtime

Custom functions require a runtime configuration in the `runtimes` array. Add a runtime object with `type` set to `"general"` and appropriate actions:

```json
{
  "extensions": [
    {
      "requirements": { /* as above */ },
      "customFunctions": { /* as above */ },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "CustomFunctionsRuntime",
                "minVersion": "1.1"
              }
            ]
          },
          "id": "CustomFunctionsRuntime",
          "type": "general",
          "code": {
            "page": "https://localhost:3000/functions.html"
          },
          "lifetime": "long",
          "actions": [
            {
              "id": "executeCustomFunctions",
              "type": "executeFunction"
            }
          ]
        }
      ]
    }
  ]
}
```

**Runtime properties explained**:

- `id`: Descriptive identifier for this runtime (for example, "CustomFunctionsRuntime").
- `type`: Must be `"general"` for custom functions.
- `code.page`: HTML page that loads your custom functions JavaScript code.
- `lifetime`: Set to `"long"` to keep the runtime alive for better performance and to enable shared runtime features.
- `actions.type`: Set to `"executeFunction"` to indicate this runtime executes functions.
- `actions.id`: Descriptive identifier for the action (for example, "executeCustomFunctions").

## Complete unified manifest example

Here's a complete unified manifest excerpt for a custom functions add-in:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
  "id": "12345678-1234-1234-1234-123456789012",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Contoso Functions",
    "full": "Contoso Custom Functions Add-in"
  },
  "description": {
    "short": "Custom functions for Excel",
    "full": "Provides custom financial and statistical functions for Excel"
  },
  "icons": {
    "outline": "https://localhost:3000/assets/outline.png",
    "color": "https://localhost:3000/assets/color.png"
  },
  "accentColor": "#D85028",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/terms"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us"
  },
  "extensions": [
    {
      "requirements": {
        "scopes": ["workbook"],
        "capabilities": [
          {
            "name": "CustomFunctionsRuntime",
            "minVersion": "1.1"
          }
        ]
      },
      "customFunctions": {
        "namespace": {
          "id": "CONTOSO",
          "name": "CONTOSO"
        },
        "metadataUrl": "https://localhost:3000/functions.json"
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "CustomFunctionsRuntime",
                "minVersion": "1.1"
              }
            ]
          },
          "id": "CustomFunctionsRuntime",
          "type": "general",
          "code": {
            "page": "https://localhost:3000/functions.html"
          },
          "lifetime": "long",
          "actions": [
            {
              "id": "executeCustomFunctions",
              "type": "executeFunction"
            }
          ]
        }
      ]
    }
  ]
}
```

## Localization with the unified manifest

To localize custom function names and descriptions, you can create locale-specific JSON metadata files and use the unified manifest's localization features.

1. Create separate JSON metadata files for each locale (for example, `functions-en.json`, `functions-de.json`).
2. Use the `localizationInfo` property in the base manifest to specify additional languages.
3. Create locale-specific override files as documented in [Localization for Office Add-ins](../develop/localization.md).

**Example localizationInfo**:

```json
{
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "de-de",
        "file": "de-de.json"
      }
    ]
  }
}
```

In your locale-specific override file (`de-de.json`), you can override the `metadataUrl` to point to the German metadata file:

```json
{
  "extensions[0].customFunctions.metadataUrl": "https://localhost:3000/functions-de.json"
}
```

For more information, see [Custom functions naming and localization](custom-functions-naming.md).

## Shared runtime configuration

Custom functions work best with a shared runtime, which allows them to:

- Interact with the task pane and share data.
- Call Office.js Excel APIs from within custom functions.
- Continue running even when the task pane is closed.

To configure a shared runtime with custom functions in the unified manifest:

1. Set `lifetime` to `"long"` in the runtime configuration (as shown in the examples above).
2. Ensure the same runtime `id` is referenced by both custom functions actions and any task pane actions.
3. Use the same `code.page` URL for both function execution and the task pane if they need to share state.

For detailed guidance on shared runtimes, see [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## Migration from add-in only manifest

If you have existing custom functions using the add-in only manifest (manifest.xml), consider these points when migrating to the unified manifest:

### Namespace compatibility

The namespace `id` and `name` in the unified manifest should match your existing configuration to avoid breaking workbooks that reference your functions.

**Add-in only manifest (XML)**:

```xml
<Namespace resid="namespace"/>
<!-- Where namespace is defined in Resources as "CONTOSO" -->
```

**Unified manifest (JSON)**:

```json
"namespace": {
  "id": "CONTOSO",
  "name": "CONTOSO"
}
```

### Metadata file reference

The way you reference the JSON metadata file changes between manifest types.

**Add-in only manifest (XML)**:

```xml
<Resources>
  <bt:Urls>
    <bt:Url id="JSON-URL" DefaultValue="https://localhost:3000/functions.json"/>
  </bt:Urls>
</Resources>
```

**Unified manifest (JSON)**:

```json
"customFunctions": {
  "metadataUrl": "https://localhost:3000/functions.json"
}
```

### Runtime configuration

The runtime configuration structure is different but serves the same purpose.

**Add-in only manifest (XML)**:

```xml
<Runtimes>
  <Runtime resid="Taskpane.Url" lifetime="long" />
</Runtimes>
```

**Unified manifest (JSON)**:

```json
"runtimes": [
  {
    "id": "CustomFunctionsRuntime",
    "type": "general",
    "code": {
      "page": "https://localhost:3000/functions.html"
    },
    "lifetime": "long",
    "actions": [
      {
        "id": "executeCustomFunctions",
        "type": "executeFunction"
      }
    ]
  }
]
```

### Function IDs must remain stable

When migrating, ensure that function IDs in your JSON metadata file remain unchanged. Changing function IDs will break existing workbooks that use your custom functions.

### Metadata JSON file unchanged

The structure of your JSON metadata file (the file referenced by `metadataUrl`) remains the same between both manifest types. You don't need to modify this file when migrating.

For complete migration guidance, see [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md).

## Troubleshooting

**Functions don't appear in Excel**:

- Verify the `metadataUrl` is accessible and returns valid JSON. Test the URL in a browser.
- Check that the namespace is configured correctly in both the manifest and metadata files.
- Ensure the CustomFunctionsRuntime requirement set is specified in the `requirements.capabilities` array.
- Clear the Office cache if testing repeatedly. See [Clear the Office cache](../testing/clear-cache.md).
- Verify that you're using Office version 2304 or later.

**Runtime errors when calling functions**:

- Verify the `code.page` URL loads correctly and includes your custom functions JavaScript.
- Check the browser console (F12 in Excel on Windows) for JavaScript errors.
- Ensure `Office.actions.associate()` correctly maps function IDs to JavaScript implementations in your code.
- Verify that the functions in your JavaScript code match the function IDs defined in your JSON metadata file.

**Shared runtime features not working**:

- Ensure `lifetime` is set to `"long"` in the runtime configuration.
- Verify that the runtime `id` matches across all references in the manifest.
- Check that both custom functions and task pane actions reference the same runtime.

For more troubleshooting guidance, see [Troubleshoot custom functions](custom-functions-troubleshooting.md).

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Manually create JSON metadata for custom functions](custom-functions-json.md)
- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Custom functions naming and localization](custom-functions-naming.md)
- [Office Add-ins with the unified manifest for Microsoft 365](../develop/unified-manifest-overview.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md)
- [Microsoft 365 extensibility schema reference - customFunctions](/microsoft-365/extensibility/schema/extension-custom-functions)
