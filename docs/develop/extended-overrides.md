---
title: Work with extended overrides of the manifest
description: Learn how to configure extensibility features with extended overrides of the manifest.
ms.topic: how-to
ms.date: 02/23/2021
ms.localizationpriority: medium
---

# Work with Extended Overrides of the manifest

Some extensibility features of Office Add-ins are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.

> [!NOTE]
> This article assumes that you're familiar with Office Add-in manifests and their role in add-ins. Please read [Office Add-ins manifest](add-in-manifests.md), if you haven't recently.

The following table specifies the extensibility features that require an extended override along with links to documentation of the feature.

| Feature | Development Instructions |
| :----- | :----- |
| Keyboard shortcuts | [Add Custom keyboard shortcuts to your Office Add-ins](../design/keyboard-shortcuts.md) |

<!-- In the following link, the "en-us" must be present or the link breaks. -->
The schema that defines the JSON format is [extended-manifest schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!TIP]
> This article is somewhat abstract. Consider reading one of the articles in the table to add clarity to the concepts.

## Tell Office where to find the JSON file

Use the manifest to tell Office where to find the JSON file. Immediately *below* (not inside) the **\<VersionOverrides\>** element in the manifest, add an [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) element. Set the `Url` attribute to the full URL of a JSON file. The following is an example of the simplest possible **\<ExtendedOverrides\>** element.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

The following is an example of a very simple extended overrides JSON file. It assigns keyboard shortcut CTRL+SHIFT+A to a function (defined elsewhere) that opens the add-in's task pane.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## Localize the extended overrides file

If your add-in supports multiple locales, you can use the `ResourceUrl` attribute of the **\<ExtendedOverrides\>** element to point Office to a file of localized resources. The following is an example.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

For more details about how to create and use the resources file, how to refer to its resources in the extended overrides file, and for additional options not discussed here, see [Localize extended overrides](localization.md#localize-extended-overrides).
