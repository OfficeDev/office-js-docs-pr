---
title: Hosts element in the manifest file
description: Specifies the Office client applications where the Office Add-in will activate.
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# Hosts element

Specifies the Office client applications where the Office Add-in will activate. Contains a collection of **Host** elements and their settings. 

## As child of VersionOverrides element

The information in this section applies *only* when the **Hosts** element is a child of a [VersionOverrides](versionoverrides.md).

This element overrides the **Hosts** element in the base manifest.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Yes   |  Describes a host and its settings. |
