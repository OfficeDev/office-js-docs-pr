---
title: Office Add-in host and platform availability
description: Supported requirement sets for Excel, Word, Outlook, PowerPoint, and OneNote.
ms.date: 12/01/2017 
---

# Office Add-in host and platform availability

To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application. 

If a table cell contains an asterisk ( * ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).  

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.

## Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Platform</th>
    <th style="width:10%">Extension points</th> 
    <th style="width:20%">API requirement sets</th> 
    <th style="width:40%"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Taskpane<br>
        - Content<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </td>
    <td>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - CompressedFile<br>
        - Settings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td>
        - Taskpane<br>
        - Content</td>
    <td>  - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Settings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td>- Taskpane<br>
        - Content<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Settings<br>
        - TextCoercion</td> 
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td>- Taskpane<br>
        - Content</td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Settings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td>- Taskpane<br>
        - Content<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

<br/>

## Outlook

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>API requirement sets</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Mail Read<br>
      - Mail Compose<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4<br>
      - Mailbox 1.5</td>
    <td>not available</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td> - Mail Read<br>
      - Mail Compose<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4</td>
    <td>not available</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td> - Mail Read<br>
      - Mail Compose<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a><br>
      - Modules</td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4<br>
      - Mailbox 1.5</td></td>
    <td>not available</td> 
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td> - Mail Read<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - Mailbox 1.4 <br>
      - Mailbox 1.5
    </td>    
    <td>not available</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td> - Mail Read<br>
      - Mail Compose<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4<br>
      - Mailbox 1.5</td>
    <td>not available</td>
  </tr>
  <tr>
    <td>Office for Android</td>
    <td> - Mail Read<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - Mailbox 1.4</td>
    <td>not available</td>
  </tr>
</table>

<br/>

## Word

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>API requirement sets</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - DocumentEvents<br>
         - TextFile<br>
         - ImageCoercion<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td> - Taskpane</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Matrix Bindings</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td> - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Matrix Bindings
    </td> 
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td> - Taskpane</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Matrix Bindings
    </td> 
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td> - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Matrix Bindings
    </td> 
  </tr>
</table>

<br/>

## PowerPoint

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>API requirement sets</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Content<br>
         - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td> - Content<br>
         - Taskpane<br>
    </td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td> - Content<br>
         - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td> - Content<br>
         - Taskpane</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
     <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td> - Content<br>
         - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
</table>

<br/>

## OneNote

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>API requirement sets</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Content<br>
         - Taskpane<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - DocumentEvents<br>
         - Settings<br>
         - TextCoercion<br>
         - HtmlCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td>Office 2016 for Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

\* = We're working on it. 

## See also

- [Office Add-ins platform overview](office-add-ins.md)
- [Common API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Add-in Commands requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [JavaScript API for Office reference](https://dev.office.com/reference/add-ins/javascript-api-for-office)

