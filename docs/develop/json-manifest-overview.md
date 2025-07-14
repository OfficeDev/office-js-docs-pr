---
title: Compare the add-in only manifest with the unified manifest for Microsoft 365
description: Get a comparison of the add-in only manifest with the unified manifest for Microsoft 365.
ms.topic: overview
ms.date: 06/24/2025
ms.localizationpriority: high
---

# Compare the add-in only manifest with the unified manifest for Microsoft 365

This article is intended to help readers who are familiar with the add-in only manifest understand the unified manifest by comparing the two. Readers should also see [Office Add-ins with the unified manifest for Microsoft 365](unified-manifest-overview.md).

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

## Schemas and general points

There is just one schema for the [unified manifest](https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/preview/DevPreview/MicrosoftTeams.schema.json), in contrast to the add-in only manifest which has a total of seven [schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

## Conceptual mapping of the unified and add-in only manifests

This section describes the unified manifest for readers who are familiar with the add-in only manifest. Some points to keep in mind:

- The unified manifest is JSON-formatted.
- JSON doesn't distinguish between attribute and element value like XML does. Typically, the JSON that maps to an XML element makes both the element value and each of the attributes a child property. The following example shows some XML markup and its JSON equivalent.
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```

- There are many places in the add-in only manifest where an element with a plural name has children with the singular version of the same name. For example, the markup to configure a custom menu includes an **\<Items\>** element which can have multiple **\<Item\>** element children. The JSON equivalent of these plural elements is a property with an array as its value. The members of the array are *anonymous* objects, not properties named "item" or "item1", "item2", etc. The following is an example.

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

### Top-level structure

The root level of the unified manifest, which roughly corresponds to the **\<OfficeApp\>** element in the add-in only manifest, is an anonymous object.

The children of **\<OfficeApp\>** are commonly divided into two notional categories. The **\<VersionOverrides\>** element is one category. The other consists of all the other children of **\<OfficeApp\>**, which are collectively referred to as the base manifest. So too, the unified manifest has a similar division. There is a top-level [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) property that roughly corresponds in its purposes and child properties to the **\<VersionOverrides\>** element. The unified manifest also has over 10 other top-level properties that collectively serve the same purposes as the base manifest of the add-in only manifest. These other properties can be thought of collectively as the base manifest of the unified manifest.

### Base manifest

The base manifest properties specify characteristics of the add-in that *any* type of extension of Microsoft 365 is expected to have. This includes Teams tabs and message extensions, not just Office add-ins. These characteristics include a public name and a unique ID. The following table shows a mapping of some critical top-level properties in the unified manifest to the XML elements in the current manifest, where the mapping principle is the *purpose* of the markup.

|JSON property|Purpose|XML elements|Comments|
|:-----|:-----|:-----|:-----|
|"$schema"| Identifies the manifest schema. | attributes of **\<OfficeApp\>** and **\<VersionOverrides\>** |*None* |
|[`"id"`](/microsoft-365/extensibility/schema/root#id)| GUID of the add-in. | **\<Id\>**|*None* |
|[`"version"`](/microsoft-365/extensibility/schema/root#version)| Version of the add-in. | **\<Version\>** |*None* |
|[`"manifestVersion"`](/microsoft-365/extensibility/schema/root#manifestversion)| Version of the manifest schema. |  attributes of **\<OfficeApp\>** |*None* |
|[`"name"`](/microsoft-365/extensibility/schema/root#name)| Public name of the add-in. | **\<DisplayName\>** |*None* |
|[`"description"`](/microsoft-365/extensibility/schema/root#description)| Public description of the add-in.  | **\<Description\>** |*None* |
|[`"accentColor"`](/microsoft-365/extensibility/schema/root#accentcolor)|*None* |*None* | This property has no equivalent in the add-in only manifest and isn't used in the unified manifest. But it must be present. |
|[`"developer"`](/microsoft-365/extensibility/schema/root#developer)| Identifies the developer of the add-in. | **\<ProviderName\>** |*None* |
|[`"localizationInfo"`](/microsoft-365/extensibility/schema/root#localizationinfo)| Configures the default locale and other supported locales. | **\<DefaultLocale\>** and **\<Override\>** |*None* |
|[`"webApplicationInfo"`](/microsoft-365/extensibility/schema/root#webApplicationInfo-property)| Identifies the add-in's web app as it is known in Microsoft Entra ID. | **\<WebApplicationInfo\>** | In the add-in only manifest, the **\<WebApplicationInfo\>** element is inside **\<VersionOverrides\>**, not the base manifest. |
|[`"authorization"`](/microsoft-365/extensibility/schema/root#authorization)| Identifies any Microsoft Graph permissions that the add-in needs. | **\<WebApplicationInfo\>** | In the add-in only manifest, the **\<WebApplicationInfo\>** element is inside **\<VersionOverrides\>**, not the base manifest. |

The **\<Hosts\>**, **\<Requirements\>**, and **\<ExtendedOverrides\>** elements are part of the base manifest in the add-in only manifest. But concepts and purposes associated with these elements are configured inside the `"extensions"` property of the unified manifest.

### `"extensions"` property

The `"extensions"` property in the unified manifest primarily represents characteristics of the add-in that wouldn't be relevant to other kinds of Microsoft 365 extensions. For example, the Office applications that the add-in extends (such as, Excel, PowerPoint, Word, and Outlook) are specified inside the `"extensions"` property, as are customizations of the Office application ribbon. The configuration purposes of the `"extensions"` property closely match those of the **\<VersionOverrides\>** element in the add-in only manifest.

> [!NOTE]
> The **\<VersionOverrides\>** section of the add-in only manifest has a "double jump" system for many string resources. Strings, including URLs, are specified and assigned an ID in the **\<Resources\>** child of **\<VersionOverrides\>**. Elements that require a string have a `resid` attribute that matches the ID of a string in the **\<Resources\>** element. The `"extensions"` property of the unified manifest simplifies things by defining strings directly as property values. There is nothing in the unified manifest that is equivalent to the **\<Resources\>** element.

The following table shows a mapping of *some* high-level child properties of the `"extensions"` property in the unified manifest to XML elements in the current manifest. Dot notation is used to reference child properties.

> [!NOTE]
> This table contains only some selected representative descendant properties of `"extensions"`. *It isn't an exhaustive list of all child properties of `"extensions"`.* For the full reference of the unified manifest, see [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

|JSON property|Purpose|XML elements|Comments|
|:-----|:-----|:-----|:-----|
| [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities) | Identifies the [requirement sets](office-versions-and-requirement-sets.md#office-requirement-sets-availability) that the add-in needs to be installable. that the add-in needs to be installable. | **\<Requirements\>** and **\<Sets\>** |*None* |
| [`"requirements.scopes"`](/microsoft-365/extensibility/schema/requirements-extension-element#scopes) | Identifies the Office applications in which the add-in can be installed. | **\<Hosts\>** |*None* |
| [`"ribbons"`](/microsoft-365/extensibility/schema/element-extensions#ribbons) | The ribbons that the add-in customizes. | **\<Hosts\>**, **ExtensionPoints**, and various **\*FormFactor** elements | The `"ribbons"` property is an array of anonymous objects that each merge the purposes of the these three elements. See [`"ribbons"` table](#ribbons-table).|
| [`"alternates"`](/microsoft-365/extensibility/schema/element-extensions#alternates) | Specifies backwards compatibility with an equivalent COM add-in, XLL, or both. | **\<EquivalentAddins\>** | See the [EquivalentAddins - See also](/javascript/api/manifest/equivalentaddins#see-also) for background information. |
| [`"runtimes"`](/microsoft-365/extensibility/schema/element-extensions#runtimes)  | Configures the [embedded runtimes](../testing/runtimes.md) that the add-in uses, including various kinds of add-ins that have little or no UI, such as custom function-only add-ins and [function commands](../design/add-in-commands.md#types-of-add-in-commands). | **\<Runtimes\>**. **\<FunctionFile\>**, and **\<ExtensionPoint\>** (of type CustomFunctions) |*None.* |
| [`"autoRunEvents"`](/microsoft-365/extensibility/schema/element-extensions#autorunevents) | Configures an event handler for a specified event. | **\<ExtensionPoint\>** (of type LaunchEvent) |*None.* |
| [`"keyboardShortcuts"`](/microsoft-365/extensibility/schema/element-extensions#keyboardshortcuts) (developer preview) | Defines custom keyboard shortcuts or key combinations to run specific actions. | **\<ExtendedOverrides\>** | *None.* |

#### `"ribbons"` table

The following table maps the child properties of the anonymous child objects in the `"ribbons"` array onto XML elements in the current manifest.

|JSON property|Purpose|XML elements|Comments|
|:-----|:-----|:-----|:-----|
| `"contexts"` | Specifies the command surfaces that the add-in customizes. | Various **\*CommandSurface** elements, such as **PrimaryCommandSurface** and **MessageReadCommandSurface** |*None.* |
| `"tabs"` | Configures custom ribbon tabs. | **\<CustomTab\>** | The names and hierarchy of the descendant properties of `"tabs"` closely match the descendants of **\<CustomTab\>**. |
| `"fixedControls"` (developer preview) | Configures and adds the button of an [integrated spam-reporting](../outlook/spam-reporting.md) add-in to the Outlook ribbon. | **\<Control\>** child element of **\<ReportPhishingCustomization\>** | *None.* |
| `"spamPreProcessingDialog"` (developer preview) | Configures the preprocessing dialog shown after the button of a spam-reporting add-in is selected from the Outlook ribbon. | **\<PreProcessingDialog\>** child element of **\<ReportPhishingCustomization\>** | *None.* |

For a full sample unified manifest, see [Sample unified manifest](unified-manifest-overview.md#sample-unified-manifest).

## Next steps

- [Build your first Outlook add-in](../quickstarts/outlook-quickstart-yo.md)
