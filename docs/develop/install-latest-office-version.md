---
title: Install the latest version of Office
description: Information about how to opt in to getting the latest builds of Office.
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# Install the latest version of Office

New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.

## Opt in to getting the latest builds of Office

- If you're a Microsoft 365 Family, Personal, or University subscriber, see [Be a Microsoft 365 Insider](https://insider.microsoft365.com).
- If you're a Microsoft 365 Apps for business customer, see [Install the First Release build for Microsoft 365 Apps for business customers](https://support.office.com/article/4dd8ba40-73c0-4468-b778-c7b744d03ead).
- If you're running Office on a Mac:
  - Start an Office application.
  - Select **Check for Updates** on the Help menu.
  - In the Microsoft AutoUpdate box, check the box to join the Microsoft 365 Insider program.

## Get the latest build of Office

1. Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).
2. Run the tool. This extracts the following two files: Setup.exe and configuration.xml.
3. Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Run the following command as an administrator:  `setup.exe /configure configuration.xml`

> [!NOTE]
> The command might take a long time to run without indicating progress.

When the installation process finishes, you will have the latest Office applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.

![A screenshot that shows product information with the Office Insiders label.](../images/office-insiders-label.png)

## Minimum Office builds for Office JavaScript API requirement sets

- [Excel JavaScript API requirement sets](/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets)
- [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [PowerPoint JavaScript API requirement sets](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets)
- [Word JavaScript API requirement sets](/javascript/api/requirement-sets/word/word-api-requirement-sets)
- [Dialog API requirement sets](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [Office Common API requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
