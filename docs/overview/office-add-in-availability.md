---
title: Office client application and platform availability for Office Add-ins
description: 'Supported requirement sets for Excel, OneNote, Outlook, PowerPoint, Project, and Word.'
ms.date: 07/13/2021
localization_priority: Priority
---

# Office client application and platform availability for Office Add-ins

To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span>Excel</span></a>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span>OneNote</span></a>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span>Outlook</span></a>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span>PowerPoint</span></a>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span>Project</span></a>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span>Word</span></a>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets. For more information about the update history of the various Office versions, check out the [See also](#see-also) section. Office Add-ins may not be supported on all services that are members of the [Office Cloud Storage Partner Program](https://developer.microsoft.com/office/cloud-storage-partner-program), which enables integrating Office on the web to work with Office documents as part of their service offering. For more information, please contact the member service.

## Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Platform</th>
    <th style="width:10%">Extension points</th>
    <th style="width:20%">API requirement sets</th>
    <th style="width:40%"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - TaskPane<br>
      - Content<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on Windows<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - TaskPane<br>
      - Content<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 on Windows<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - Content<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Windows<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - Content
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 on Windows<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - Content
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on iPad<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - TaskPane<br>
      - Content
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on Mac<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - TaskPane<br>
      - Content<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 on Mac<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - Content<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Mac<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - Content
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; - Added with post-release updates.*

## Custom Functions (Excel only)

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th>
    <th>API requirement sets</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office on Windows<br>(connected to a Microsoft 365 subscription)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office on Mac<br>(connected to a Microsoft 365 subscription)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
</table>

## Outlook

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th>
    <th>API requirement sets</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office on the web<br>(modern)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office on the web<br>(classic)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office on Windows<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office 2019 on Windows<br>(one-time purchase)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office 2016 on Windows<br>(one-time purchase)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office 2013 on Windows<br>(one-time purchase)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office on iOS<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office on Mac<br>(current UI,<br>connected to a Microsoft 365 subscription)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office on Mac<br>(new UI (preview)<sup>3</sup>,<br>connected to a Microsoft 365 subscription)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office 2019 on Mac<br>(one-time purchase)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office 2016 on Mac<br>(one-time purchase)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Not available</td>
  </tr>
  <tr>
    <td>Office on Android<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </td>
    <td>Not available</td>
  </tr>
</table>

> [!NOTE]
> <sup>1</sup> To require Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`. Declaring it in your add-in's manifest isn't supported. You can also determine if the API is supported by checking that it's not `undefined`. For further details, see [Using APIs from later requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).
>
> <sup>2</sup> Added with post-release updates.
>
> <sup>3</sup> Support for the new Mac UI (preview) is available from Outlook version 16.38.506. For more information, see the [Add-in support in Outlook on new Mac UI](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) section.

> [!IMPORTANT]
> Client support for a requirement set may be restricted by Exchange server support. See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.

<br/>

## Word

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th>
    <th>API requirement sets</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office on Windows<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 on Windows<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Windows<br>(one-time purchase)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 on Windows<br>(one-time purchase)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office on iPad<br>(connected to a Microsoft 365 subscription)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office on Mac<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 on Mac<br>(one-time purchase)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Mac<br>(one-time purchase)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
</table>

*&ast; - Added with post-release updates.*

<br/>

## PowerPoint

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th>
    <th>API requirement sets</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - Content<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on Windows<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - Content<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 on Windows<br>(one-time purchase)</td>
    <td>
      - Content<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Windows<br>(one-time purchase)</td>
    <td>
      - Content<br>
      - TaskPane
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 on Windows<br>(one-time purchase)</td>
    <td>
      - Content<br>
      - TaskPane
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on iPad<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - Content<br>
      - TaskPane
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office on Mac<br>(connected to a Microsoft 365 subscription)</td>
    <td>
      - Content<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 on Mac<br>(one-time purchase)</td>
    <td>
      - Content<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Mac<br>(one-time purchase)</td>
    <td>
      - Content<br>
      - TaskPane
    </td>
    <td>
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; - Added with post-release updates.*

<br/>

## OneNote

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th>
    <th>API requirement sets</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office on the web</td>
    <td>
      - Content<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## Project

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th>
    <th>API requirement sets</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></th>
  </tr>
  <tr>
    <td>Office 2019 on Windows<br>(one-time purchase)</td>
    <td>- TaskPane</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 on Windows<br>(one-time purchase)</td>
    <td>- TaskPane</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 on Windows<br>(one-time purchase)</td>
    <td>- TaskPane</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## See also

- [Office Add-ins platform overview](office-add-ins.md)
- [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md)
- [Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Add-in Commands requirement sets](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [API Reference documentation](../reference/javascript-api-for-office.md)
- [Update history for Microsoft 365 Apps](/officeupdates/update-history-office365-proplus-by-date)
- [Office 2016 and 2019 update history (Click-To-Run)](/officeupdates/update-history-office-2019)
- [Office 2013 update history (Click-To-Run)](/officeupdates/update-history-office-2013)
- [Office 2010, 2013, and 2016 update history (MSI)](/officeupdates/office-updates-msi)
- [Outlook 2010, 2013, and 2016 update history (MSI)](/officeupdates/outlook-updates-msi)
- [Update history for Office for Mac](/officeupdates/update-history-office-for-mac)
- [Develop Office Add-ins](../develop/develop-overview.md)
