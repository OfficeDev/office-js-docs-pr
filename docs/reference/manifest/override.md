---
title: Override element in the manifest file
description: The Override element enables you to specify the value of a setting depending on a specified condition.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# Override element

Provides a way to override the value of a manifest setting depending on a specified condition. There are three kinds of conditions:

- An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.
- A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.
- The source is different from the default `Runtime`, called **RuntimeOverride**.

An **Override** element that is inside of a **Runtime** element must be of type **RuntimeOverride**.

There is no `overrideType` attribute for the **Override** element. The difference is determined by the parent element and the parent element's type. An **Override** element that is inside of a **Token** element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**. An **Override** element inside any other parent element, or inside an **Override** element of type `LocaleToken`, must be of type **LocaleTokenOverride**. For more information about the use of this element when it's a child of a **Token** element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

Each type is described in separate sections later in this article.

## Override element for `LocaleToken`

An **Override** element expresses a conditional and can be read as an "If ... then ..." statement. If the **Override** element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent. For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**Add-in type:** Content, Task pane, Mail

### Syntax

```XML
<Override Locale="string" Value="string"></Override>
```

### Contained in

|Element|
|:-----|
|[CitationText](citationtext.md)|
|[Description](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|
|[Token](token.md)|

### Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|Locale|string|required|Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.|
|Value|string|required|Specifies value of the setting expressed for the specified locale.|

### Examples

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### See also

- [Localization for Office Add-ins](../../develop/localization.md)
- [Keyboard shortcuts](../../design/keyboard-shortcuts.md)

## Override element for `RequirementToken`

An **Override** element expresses a conditional and can be read as an "If ... then ..." statement. If the **Override** element is of type **RequirementTokenOverride**, then the child **Requirements** element expresses the condition, and the `Value` attribute is the consequent. For example, the first **Override** in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**Add-in type:** Task pane

### Syntax

```XML
<Override Value="string" />
```

### Contained in

|Element|
|:-----|
|[Token](token.md)|

### Must contain

|Element|Content|Mail|TaskPane|
|:-----|:-----|:-----|:-----|
|[Requirements](requirements.md)|||x|

### Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|Value|string|required|Value of the grandparent token when the condition is satisfied.|

### Example

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify which Office versions and platforms can host your add-in](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)
- [Keyboard shortcuts](../../design/keyboard-shortcuts.md)

## Override element for `Runtime`

> [!IMPORTANT]
> Support for this element was introduced in [Mailbox requirement set 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) with the [event-based activation feature](../../outlook/autolaunch.md). See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

An **Override** element expresses a conditional and can be read as an "If ... then ..." statement. If the **Override** element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent. For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**Add-in type:** Mail

### Syntax

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### Contained in

- [Runtime](runtime.md)

### Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|**type**|string|Yes|Specifies the language for this override. At present, `"javascript"` is the only supported option.|
|**resid**|string|Yes|Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`. The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.|

### Examples

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### See also

- [Runtime](runtime.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)
