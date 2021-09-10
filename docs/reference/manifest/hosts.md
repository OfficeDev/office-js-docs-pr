---
title: Hosts element in the manifest file
description: Specifies the Office client application where the Office Add-in will activate.
ms.date: 10/09/2018
ms.localizationpriority: medium
---

# Hosts element

Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings. 

When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest. 

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Yes   |  Describes a host and its settings. |
