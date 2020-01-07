---
title: Install the latest version of Office
description: Information about how to opt in to getting the latest builds of Office.
ms.date: 01/07/2020
localization_priority: Priority
---

# Install the latest version of Office

New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.

## Opt in to getting the latest builds

To opt in to getting the latest builds of Office:

- If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).
- If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- If you're running Office on a Mac:
  - Start an Office application.
  - Select **Check for Updates** on the Help menu.
  - In the Microsoft AutoUpdate box, check the box to join the Office Insider program.

## Get the latest build

To get the latest build of Office:

1. Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).
2. Run the tool. This extracts the following two files: Setup.exe and configuration.xml.
3. Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Run the following command as an administrator:  `setup.exe /configure configuration.xml`

> [!NOTE]
> The command might take a long time to run without indicating progress.

When the installation process finishes, you will have the latest Office applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.

![A screenshot that shows product information with the Office Insiders label](../images/officeinsiderslabel.png)

## Minimum Office builds for Office JavaScript API requirement sets

For information about the minimum product builds for each platform for the API requirement sets, see the following:

- [Excel JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [OneNote JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [Outlook JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
- [PowerPoint JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets)
- [Word JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [Dialog API requirement sets](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [Office Common API requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
