---
title: Specify Office hosts and API requirements with the add-in only manifest
description: Learn how to specify Office applications and API requirements for your add-in to work as expected.
ms.topic: best-practice
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Specify Office applications and API requirements with the add-in only manifest

> [!NOTE]
> For information about specifying requirements with the [unified manifest for Microsoft 365](unified-manifest-overview.md), see [Specify Office hosts and API requirements with the unified manifest](specify-office-hosts-and-api-requirements-unified.md).

Your Office Add-in might depend on a specific Office application (also called an Office host) or on specific members of the Office JavaScript Library (office.js). For example, your add-in might:

- Run in a single Office application (for example, Word or Excel), or several applications.
- Make use of Office JavaScript APIs that are only available in some versions of Office. For example, the volume-licensed perpetual version of Excel 2016 doesn't support all Excel-related APIs in the Office JavaScript library.

In these situations, you need to ensure that your add-in is never installed on Office applications or Office versions in which it cannot run.

There are also scenarios in which you want to control which features of your add-in are visible to users based on their Office application and Office version. Two examples are:

- Your add-in has features that are useful in both Word and PowerPoint, such as text manipulation, but it has some additional features that only make sense in PowerPoint, such as slide management features. You need to hide the PowerPoint-only features when the add-in is running in Word.
- Your add-in has a feature that requires an Office JavaScript API method that is supported in some versions of an Office application, such as Microsoft 365 subscription Excel, but is not supported in others, such as volume-licensed perpetual Excel 2016. But your add-in has other features that require only Office JavaScript API methods that *are* supported in volume-licensed perpetual Excel 2016. In this scenario, you need the add-in to be installable on that version of Excel 2016, but the feature that requires the unsupported method should be hidden from those users.

This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.

> [!NOTE]
> For a high-level view of where Office Add-ins are currently supported, see the [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets) page.

> [!TIP]
> Many of the tasks described in this article are done for you, in whole or in part, when you create your add-in project with a tool, such as the [Yeoman generator for Office Add-ins](yeoman-generator-overview.md) or one of the Office Add-in templates in Visual Studio. In such cases, please interpret the task as meaning that you should verify that it has been done.

## Use the latest Office JavaScript API library

Your add-in should load the most current version of the Office JavaScript API library from the content delivery network (CDN). To do this, be sure you have the following `<script>` tag in the first HTML file your add-in opens. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## Specify which Office applications can host your add-in

By default, an add-in is installable in all Office applications supported by the specified add-in type (that is, Mail, Task pane, or Content). For example, a task pane add-in is installable by default on Access, Excel, OneNote, PowerPoint, Project, and Word.

To ensure that your add-in is installable in only a subset of Office applications, use the [Hosts](/javascript/api/manifest/hosts) and [Host](/javascript/api/manifest/host) elements in the add-in only manifest.

For example, the following `<Hosts>` and `<Host>` declaration specifies that the add-in can install on any release of Excel, which includes Excel on the web, Windows, and iPad, but can't be installed on any other Office application.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

The `<Hosts>` element can contain one or more `<Host>` elements. There should be a separate `<Host>` element for each Office application on which the add-in should be installable. The `Name` attribute is required and can be set to one of the following values.

| Name          | Office client applications                     | Available add-in types |
|:--------------|:-----------------------------------------------|:-----------------------|
| Document      | Word on the web, Windows, Mac, iPad            | Task pane              |
| Mailbox       | Outlook on the web, Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), Mac, Android, iOS | Mail |
| Notebook      | OneNote on the web                             | Task pane, Content     |
| Presentation  | PowerPoint on the web, Windows, Mac, iPad      | Task pane, Content     |
| Project       | Project on Windows                             | Task pane              |
| Workbook      | Excel on the web, Windows, Mac, iPad           | Task pane, Content     |
| Database      | Access (obsolete)                              | Task pane              |

> [!NOTE]
> Office applications are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You usually can't specify which platform can be used to run your add-in. For example, if you specify `Workbook`, both Excel on the web and on Windows can be used to run your add-in. However, if you specify `Mailbox`, your add-in won't run on Outlook mobile clients unless you define the [mobile extension point](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface).

> [!NOTE]
> It isn't possible for an add-in only manifest to apply to more than one type: Mail, Task pane, or Content. This means that if you want your add-in to be installable on Outlook and on one of the other Office applications, you must create *two* add-ins, one with a Mail type manifest and the other with a Task pane or Content type manifest.

## Specify which Office versions and platforms can host your add-in

You can't explicitly specify the Office versions and builds or the platforms on which your add-in should be installable, and you wouldn't want to because you would have to revise your manifest whenever support for the add-in features that your add-in uses is extended to a new version or platform. Instead, specify in the manifest the APIs that your add-in needs. Office prevents the add-in from being installed on combinations of Office version and platform that don't support the APIs and ensures that the add-in won't appear in **My Add-ins**.

> [!IMPORTANT]
> Only use the base manifest to specify the API members that your add-in must have to be of any significant value at all. If your add-in uses an API for some features but has other useful features that don't require the API, you should design the add-in so that it's installable on platform and Office version combinations that don't support the API but provides a diminished experience on those combinations. For more information, see [Design for alternate experiences](#design-for-alternate-experiences).

### Requirement sets

To simplify the process of specifying the APIs that your add-in needs, Office groups most APIs together in [requirement sets](office-versions-and-requirement-sets.md). The APIs in the [Common API Object Model](understanding-the-javascript-api-for-office.md#api-models) are grouped by the development feature that they support. For example, all the APIs connected to table bindings are in the requirement set called "TableBindings 1.1". The APIs in the [Application specific object models](understanding-the-javascript-api-for-office.md#api-models) are grouped by when they were released for use in production add-ins.

Requirement sets are versioned. For example, the APIs that support [Dialog Boxes](../develop/dialog-api-in-office-add-ins.md) are in the requirement set DialogApi 1.1. When additional APIs that enable messaging from a task pane to a dialog were released, they were grouped into DialogApi 1.2, along with all the APIs in DialogApi 1.1. *Each version of a requirement set is a superset of all earlier versions.*

Requirement set support varies by Office application, the version of the Office application, and the platform on which it is running. For example, ExcelApi 1.17 isn't supported on volume-licensed perpetual versions of Office before Office 2024 but ExcelApi 1.14 is supported back to Office 2021. You want your add-in to be installable on every combination of platform and Office version that supports the APIs that it uses, so you should always specify in the manifest the *minimum* version of each requirement set that your add-in requires. Details about how to do this are later in this article.

> [!TIP]
> For more information about requirement set versioning, see [Office requirement sets availability](office-versions-and-requirement-sets.md#office-requirement-sets-availability), and for the complete lists of requirement sets and information about the APIs in each, start with [Office Add-in requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets). The reference topics for most Office.js APIs also specify the requirement set they belong to (if any).

> [!NOTE]
> Some requirement sets also have manifest elements associated with them. See [Specifying requirements in a VersionOverrides element](#specify-requirements-in-a-versionoverrides-element) for information about when this fact is relevant to your add-in design.

### Requirements element

Use the [Requirements](/javascript/api/manifest/requirements) element and its child element [Sets](/javascript/api/manifest/sets) to specify the minimum requirement sets that must be supported by the Office application to install your add-in.

All APIs in the application specific models are in requirement sets, but some of those in the Common API model are not. Use the [Methods](/javascript/api/manifest/methods) to specify the setless API members that your add-in requires. You can't use the `<Methods>` element with Outlook add-ins.

If the Office application or platform doesn't support the requirement sets or API members specified in the `<Requirements>` element, the add-in won't run in that application or platform, and won't display in **My Add-ins**.

> [!NOTE]
> The `<Requirements>` element is optional for all add-ins, except for Outlook add-ins. When the `xsi:type` attribute of the root `OfficeApp` element is `MailApp`, there must be a `<Requirements>` element that specifies the minimum version of the Mailbox requirement set that the add-in requires. For more information, see [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

The following code example shows how to configure an add-in that is installable in all Office applications that support the following:

- `TableBindings` requirement set, which has a minimum version of "1.1".
- `OOXML` requirement set, which has a minimum version of "1.1".
- `Document.getSelectedDataAsync` method.

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```

Note the following about this example.

- The `<Requirements>` element contains the `<Sets>` and `<Methods>` child elements.
- The `<Sets>` element can contain one or more `<Set>` elements. `DefaultMinVersion` specifies the default `MinVersion` value of all child `<Set>` elements.
- A [Set](/javascript/api/manifest/set) element specifies a requirement set that the Office application must support to make the add-in installable. The `Name` attribute specifies the name of the requirement set. The `MinVersion` specifies the minimum version of the requirement set. `MinVersion` overrides the value of the `DefaultMinVersion` attribute in the parent `<Sets>`.
- The `<Methods>` element can contain one or more [Method](/javascript/api/manifest/method) elements. You can't use the `<Methods>` element with Outlook add-ins.
- The `<Method>` element specifies an individual method that the Office application must support to make the add-in installable. The `Name` attribute is required and specifies the name of the method qualified with its parent object.

## Design for alternate experiences

The extensibility features that the Office Add-in platform provides can be usefully divided into three kinds:

- Extensibility features that are available immediately after the add-in is installed. You can make use of this kind of feature by configuring a [VersionOverrides](/javascript/api/manifest/versionoverrides) element in the manifest. An example of this kind of feature is [Add-in Commands](../design/add-in-commands.md), which are custom ribbon buttons and menus.
- Extensibility features that are available only when the add-in is running and that are implemented with Office.js JavaScript APIs; for example, [Dialog Boxes](../develop/dialog-api-in-office-add-ins.md).
- Extensibility features that are available only at runtime but are implemented with a combination of Office.js JavaScript and configuration in a `<VersionOverrides>` element. Examples of these are [Excel custom functions](../excel/custom-functions-overview.md), [single sign-on](sso-in-office-add-ins.md), and [custom contextual tabs](../design/contextual-tabs.md).

If your add-in uses a specific extensibility feature for some of its functionality but has other useful functionality that doesn't require the extensibility feature, you should design the add-in so that it's installable on platform and Office version combinations that don't support the extensibility feature. It can provide a valuable, albeit diminished, experience on those combinations.

You implement this design differently depending on how the extensibility feature is implemented:

- For features implemented entirely with JavaScript, see [Check for API availability at runtime](specify-api-requirements-runtime.md).
- For features that require you to configure a `<VersionOverrides>` element, see [Specifying requirements in a VersionOverrides element](#specify-requirements-in-a-versionoverrides-element).

### Specify requirements in a VersionOverrides element

The [VersionOverrides](/javascript/api/manifest/versionoverrides) element was added to the manifest schema primarily, but not exclusively, to support features that must be available immediately after an add-in is installed, such as add-in commands (custom ribbon buttons and menus). Office must know about these features when it parses the add-in manifest.

Suppose your add-in uses one of these features, but the add-in is valuable, and should be installable, even on Office versions that don't support the feature. In this scenario, identify the feature using a [Requirements](/javascript/api/manifest/requirements) element (and its child [Sets](/javascript/api/manifest/sets) and [Methods](/javascript/api/manifest/methods) elements) that you include as a child of the `<VersionOverrides>` element itself instead of as a child of the base `OfficeApp` element. The effect of doing this is that Office will allow the add-in to be installed, but Office will ignore certain of the child elements of the `<VersionOverrides>` element on Office versions where the feature isn't supported.

Specifically, the child elements of the `<VersionOverrides>` that override elements in the base manifest, such as a `<Hosts>` element, are ignored and the corresponding elements of the base manifest are used instead. However, there can be child elements in a `<VersionOverrides>` that actually implement additional features rather than override settings in the base manifest. Two examples are the `WebApplicationInfo` and `EquivalentAddins`. These parts of the `<VersionOverrides>` will *not* be ignored, assuming the platform and version of Office support the corresponding feature.  

For information about the descendent elements of the `<Requirements>` element, see [Requirements element](#requirements-element) earlier in this article.

The following is an example.

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> If your add-in includes [add-in commands](../design/add-in-commands.md), use great care before including a `<Requirements>` element in a `<VersionOverrides>`. On platform and version combinations that don't support the requirement, *none* of the add-in commands will be installed, *even those that invoke functionality that doesn't need the requirement*. Consider, for example, an add-in that has two custom ribbon buttons. One of them calls Office JavaScript APIs that are available in requirement set **ExcelApi 1.4** (and later). The other calls APIs that are only available in **ExcelApi 1.9** (and later). If you put a requirement for **ExcelApi 1.9** in the `<VersionOverrides>`, then when 1.9 isn't supported, *neither* button will appear on the ribbon. A better strategy in this scenario would be to use the technique described in [Check for API availability at runtime](specify-api-requirements-runtime.md). The code invoked by the second button first uses `isSetSupported` to check for support of **ExcelApi 1.9**. If it isn't supported, the code gives the user a message saying that this feature of the add-in isn't available on their version of Office.

> [!TIP]
> There's no point to repeating a `<Requirement>` element in a `<VersionOverrides>` that already appears in the base manifest. If the requirement is specified in the base manifest, then the add-in can't install where the requirement isn't supported so Office doesn't even parse the `<VersionOverrides>` element.

## See also

- [Office Add-ins manifest](add-in-manifests.md)
- [Office Add-in requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
