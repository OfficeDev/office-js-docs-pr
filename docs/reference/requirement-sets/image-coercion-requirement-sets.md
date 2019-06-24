---
title: Image Coercion requirement sets
description: ''
ms.date: 06/24/2019
ms.prod: non-product-specific
localization_priority: Normal
---

# Image Coercion requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office on Windows |  Office on iPad |  Office on Mac | Office on the web  |  Office Online Server  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| ImageCoercion 1.1  | <NEED VERSION> | <NEED VERSION> | <NEED VERSION> | <NEED VERSION> | <NEED VERSION> |
| ImageCoercion 1.2  | <NEED VERSION> | <NEED VERSION> | <NEED VERSION> | <NEED VERSION> | <NEED VERSION> |

## ImageCoercion 1.1

ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method. The following hosts are supported:

- Excel 2013 and later on Windows
- Excel 2016 and later on Mac
- Excel on the web
- Excel on iPad
- OneNote on the web
- PowerPoint 2013 and later on Windows
- PowerPoint 2016 and later on Mac
- PowerPoint on the web
- PowerPoint on iPad
- Word 2013 and later on Windows
- Word 2016 and later on Mac
- Word on the web
- Word on iPad

## ImageCoercion 1.2

ImageCoercion 1.2 enables conversion to an SVG (Office.CoercionType.XmlSvg) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method. The following hosts are supported:

- Excel on Windows (connected to and Office 365 subscription)
- Excel on Mac (connected to and Office 365 subscription)
- Excel on the web
- Excel on iPad
- PowerPoint on Windows (connected to and Office 365 subscription)
- PowerPoint on Mac (connected to and Office 365 subscription)
- PowerPoint on the web
- PowerPoint on iPad
- Word on Windows (connected to and Office 365 subscription)
- Word on Mac (connected to and Office 365 subscription)
- Word on the web
- Word on iPad


## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
