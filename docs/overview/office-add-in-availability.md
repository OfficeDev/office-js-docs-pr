# Office Add-in host and platform availability

See the following tables to find the Office application and platform that you want to work with to learn about the [supported requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets). 

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.

If a table cell is empty, that means we're working on it.

## Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Platform</th>
    <th style="width:10%">Extension points</th> 
    <th style="width:20%">APIs</th> 
    <th style="width:40%"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Shared APIs</b></a></th> 
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
    <td>  - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1 (Build 15.0.4855.1000+)</a></td>
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
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands (Build 16.0.6868.1000+)</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1 (Build 4266.1001+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2 (Build 6741.2088+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3 (Build 7369.2055+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4 (Build 7870.2024+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1 (Build 6741.0000+)</a></td>
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
    <td>Office for iPad</td>
    <td>- Taskpane<br>
        - Content</td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1 (Build 1.19+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2 (Build 1.21+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3 (Build 1.27+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1 (Build 1.22+)</a></td>
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
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands (Build 15.33+)</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1 (Build 15.20+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2 (Build 15.22+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3 (Build 15.27+)</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1 (Build 15.20+)</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office for iPhone</td>
    <td> </td>
    <td> </td>
    <td> </td>
  </tr>
  <tr>
    <td>Office for Android</td>
    <td> </td>
    <td> </td>
    <td> </td>
  </tr>
  <tr>
    <td>Office Mobile for Windows 10</td>
    <td> </td>
    <td> </td>
    <td> </td>
  </tr>
</table>

## Outlook

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>APIs</th> 
    <th>Shared APIs</th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td>
    - Mail Read<br>
    - Mail Compose<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>
    - Mailbox 1.0<br>
    - Mailbox 1.1<br>
    - Mailbox 1.2<br>
    - Mailbox 1.3<br>
    - Mailbox 1.4<br>
    - Mailbox 1.5</td>
    <td> </td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    - Mail Read<br>
    - Mail Compose<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>
    - Mailbox 1.0<br>
    - Mailbox 1.1<br>
    - Mailbox 1.2<br>
    - Mailbox 1.3<br>
    - Mailbox 1.4</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    - Mail Read<br>
    - Mail Compose<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands (Build 16.0+)</a></td>
    <td>    
    - Mailbox 1.0<br>
    - Mailbox 1.1<br>
    - Mailbox 1.2<br>
    - Mailbox 1.3<br>
    - Mailbox 1.4<br>
    - Mailbox 1.5</td></td>
    <td> </td> 
  </tr>
  <tr>
    <td>Office for iPad</td>
    - Mail Read<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>Mailbox 1.4</td>
    <td> </td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    - Mail Read<br>
    - Mail Compose<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands (Build 15.32+)</a></td>
    <td>
    - Mailbox 1.0<br>
    - Mailbox 1.1<br>
    - Mailbox 1.2<br>
    - Mailbox 1.3<br>
    - Mailbox 1.4<br>
    - Mailbox 1.5</td>
    <td> </td>
  </tr>
  <tr>
    <td>Office for iPhone</td>
    - Mail Read<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>Mailbox 1.4</td>
    <td> </td>
  </tr>
  <tr>
    <td>Office for Android</td>
    - Mail Read<br>
    - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></td>
    <td>Mailbox 1.4</td>
    <td> </td>
  </tr>
  <tr>
    <td>Office Mobile for Windows 10</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
</table>

## Word

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>APIs</th> 
    <th>Shared APIs</th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td> 
  </tr>
  <tr>
    <td>Office for iPad</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office for iPhone</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office for Android</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office Mobile for Windows 10</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
</table>

## PowerPoint

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>APIs</th> 
    <th>Shared APIs</th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td> 
  </tr>
  <tr>
    <td>Office for iPad</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office for iPhone</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office for Android</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office Mobile for Windows 10</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
</table>

## OneNote

<table style="width:80%">
  <tr>
    <th>Platform</th>
    <th>Extension points</th> 
    <th>APIs</th> 
    <th>Shared APIs</th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td> 
  </tr>
  <tr>
    <td>Office for iPad</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office for iPhone</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office for Android</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
  <tr>
    <td>Office Mobile for Windows 10</td>
    <td>Description</td>
    <td>Description</td>
    <td>Description</td>
  </tr>
</table>


> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.

## Additional resources

- [Office Add-ins platform overview](office-add-ins.md)
- [Office common API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [JavaScript API for Office reference](https://dev.office.com/reference/add-ins/javascript-api-for-office)

