# Install the latest version of Office 2016

New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office. To opt in the latest builds of Office 2016: 

- If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/en-us/office-insider).
- If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/en-us/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead?ui=en-US&rs=en-US&ad=US).

To get the latest build: 

1. Download the [Office 2016 Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117). 
2. Run the tool. This extracts the following two files: Setup.exe and configuration.xml.
3. Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Run the following command as an administrator:  `setup.exe /configure configuration.xml` 

When the installation process finishes, you will have the latest Office 2016 applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.

![A screenshot that shows product information with the Office Insiders label](../../images/officeinsider.PNG)
