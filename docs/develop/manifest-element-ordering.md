---
title: How to find the proper order of manifest elements
description: Learn how to find the correct order in which to place child elements in a parent element.
ms.date: 11/16/2018
---

# How to find the proper order of manifest elements

The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.

The required ordering is specified in the XSD files in the [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) folder. The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.

For example, In the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order. If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element. Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.

> [!NOTE]
> The [Office Add-in Validator](/office/dev/add-ins/testing/troubleshoot-manifest?branch=manifest-element-ordering#validate-your-manifest-with-the-office-add-in-validator) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent. The error says the child element is is not a valid child of the parent element. If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.

To find the correct order for the child elements of a given parent element, take the following steps. (This is a simplified process, as XSD files are quite complex. Fully parsing XSD files is out of the scope of this document.)

1. Open the subfolder under [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) for the type of add-in that you are creating. 
2. Open the XSD file where the parent element is defined as a complex type. If you don't know which file has the definition, you may have to do step 3 on multiple files until you find it.
3. Search for `<xs:complexType name="PARENT_ELEMENT">`, where PARENT_ELEMENT is the name of the parent element.
4. Inside the definition for the PARENT_ELEMENT, there is (usually) an element called `<xs:sequence>`. The following is the definition for `<SuperTip>` from [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd).

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

The `<xs:sequence>` lists the possible child elements, *in the order in which they must appear*. This does *not* mean all of them are mandatory. If the `minOccurs` value for a child element is **0**, then the child element is optional. *But if it is present, it must be in the order specified by the `<xs:sequence>` element*.

If there is no `<xs:sequence>` element, or there *is* but the child element is not listed (even though the reference documentation for the child element indicates that it *is* valid for the parent); then the parent element's complex type definition has been extended with additional child elements somewhere else in the XSD file. For example, the definition for the `OfficeApp` complex type does not list `Requirements` as a possible child. But later in the file (within the definition for the `TaskPaneApp` complex type), the definition of `OfficeApp` is extended and `Requirements` is added as an additional valid child.

To find the extended definitions follow these steps:

1. Starting at the top of the file, search for `<xs:extension base="PARENT_ELEMENT">`, where PARENT_ELEMENT is the name of the parent element. There may be more than one extension.
2. Find the extension that is relevant to the context in which you are working. For example, the `OfficeApp` complex type is extended within the `ContentApp` and `MailApp` complex types as well as within the `TaskPaneApp` complex type.

Each `<xs:extension base="PARENT_ELEMENT">` in the file has its own `<xs:sequence>` that lists additional valid child elements for the parent. Child elements on an extended list must always be *after* the child elements in the original list in the parent's complex type definition.

## See also

- [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md)
