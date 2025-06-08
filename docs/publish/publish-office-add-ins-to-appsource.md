---
title: Publish your Office Add-in to Microsoft AppSource
description: Learn how to publish your Office Add-in to Microsoft AppSource and install the add-in with a Windows app or COM/VSTO add-in.
ms.topic: concept-article
ms.date: 06/06/2025
CustomerIntent: As a developer, I want to publish my Office Add-in to Microsoft AppSource so that customers can deploy and use my new add-in.
---

# Publish your Office Add-in to Microsoft AppSource

Publish your Office Add-in to [Microsoft AppSource](https://appsource.microsoft.com/) to make it widely available to customers and businesses. Microsoft AppSource is an online store that contains thousands of business applications and services built by industry-leading software providers. When you publish your add-in to Microsoft AppSource, you also make it available in the in-product experience within Office.

## The publishing process

Before you proceed:

- Have a [Partner Center account](/partner-center/marketplace-offers/open-a-developer-account).
- Ensure that your add-in adheres to the applicable [AppSource validation policies](/legal/marketplace/certification-policies).
- Confirm that you're [ready to publish](/partner-center/marketplace-offers/checklist).

When you're ready to include your solution in Microsoft AppSource and within Office, submit it to Partner Center. Then, it goes through an approval and certification process. For complete details, see [Make your solutions available in Microsoft AppSource and within Office](/partner-center/marketplace/submit-to-appsource-via-partner-center).

When your add-in is available in AppSource, there are two further steps you can take to make it more widely installed. 

- [Provide an installation link](#provide-an-installation-link)
- [Include the add-in in the installation of a Windows app or a COM or VSTO add-in](#include-the-add-in-in-the-installation-of-a-windows-app-or-comvsto-add-in)

### Provide an installation link

After you publish to Microsoft AppSource, you can create an installation link to help customers discover and install your add-in. The installation link provides a "click and run" experience. Put the link on your website, social media, or anywhere you think helps your customers discover your add-in. When users select the link, the Office application opens with a new document and your add-in is installed without the need for users to search for it in Microsoft AppSource and install it manually.

> [!NOTE]
> Installation links can only be created for Excel, PowerPoint, and Word add-ins.

There are two styles of links depending on whether you want the Office application to open in Office on the web or on desktop Office.

- [Open in Office on the web](#open-in-office-on-the-web)
- [Open in Office on Windows or Mac](#open-in-office-on-windows-or-mac)

> [!NOTE]
> The style of link determines only the platform that opens, not the platforms where the add-in is installed. The add-in is installed on every platform that supports all of the requirement sets that are specified in its manifest. See [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).

#### Open in Office on the web

To create the link, use the following URL template as a reference.

`https://go.microsoft.com/fwlink/?linkid={{linkId}}&templateid={{addInId}}&templatetitle={{addInName}}`

Change the placeholders in the URL as follows.

- **linkId**: Specifies which web endpoint to use when opening the new document.

  - For Word on the web: `2261098`
  - For Excel on the web: `2261819`
  - For PowerPoint on the web: `2261820`

- **addInId**: The ID of your add-in as listed in AppSource.
- **addInName**: The full title of your add-in. This must be URL-encoded.

For example, the following is an installation link for [Script Lab](https://appsource.microsoft.com/product/office/wa104380862).

[https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project](https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project)

The following parameter values are used for the Script Lab installation link.

- **linkid:** The value `2261819` specifies the Excel endpoint. Script Lab supports Word, Excel, and PowerPoint, so this value can be changed to support different endpoints.
- **templateid:** The value `WA104380862` is the Microsoft AppSource ID for Script Lab.
- **templatetitle:** The value `Script%20Lab,%20a%20Microsoft%20Garage%20project` is the URL-encoded value of the title.

#### Open in Office on Windows or Mac

To create the link, use the following URL template as a reference.

`{{appName}}:https://api.addins.store.office.com/addinstemplate/{{language}}/{{correlationID}}/{{addinId}}/none/{{addinName}}.{{fileFormat}}?omexsrctype=1&isexternallink=1`

Change the placeholders in the URL as follows.

- **appName:** This parameter specifies which Office application is opened when users click the link.

    - For Word: `ms-word`
    - For Excel: `ms-excel`
    - For PowerPoint: `ms-powerpoint`

- **language:** This is the language of the add-in. For example: `de-DE`, `ja-JP`, or `zh-CN`.
- **correlationID:** A GUID for diagnostic purposes. For example, "7bf846ec-905a-5edd-b162-83498f9a8674". Use a GUID generation tool. The diagnostic purpose is defeated if multiple users have the same GUID, so we recommend you to generate it at runtime and make it different per click.
- **addInId:** The ID of your add-in as listed in AppSource. *This parameter is case-sensitive. Use uppercase letters.*
- **addInName:** The full name of your add-in. This must be URL-encoded.
- **fileFormat:** Specifies the format of the file that is created.

    - For Word: `docx`
    - For Excel: `xlsx`
    - For PowerPoint: `pptx`

For example, the following is an installation link for Script Lab in Word on Windows with en-US as the language.

`ms-word:https://api.addins.store.office.com/addinstemplate/en-US/228a829b-69d7-45f4-a338-c6aba330ec7e/WA104380862/none/Script%20Lab,%20a%20Microsoft%20Garage%20project.docx?omexsrctype=1&isexternallink=1`

### Include the add-in in the installation of a Windows app or COM/VSTO add-in

When you have a Windows app or a COM or VSTO add-in whose functions overlap with your Office Web Add-in, consider including the web add-in in the installation (or an upgrade) of the Windows app or COM/VSTO add-in. (This installation option is supported only for Excel, PowerPoint, and Word add-ins.) The process for doing this depends on whether you are a [certified Microsoft 365 developer](/microsoft-365-app-certification/docs/certification). For more information, see [Microsoft 365 App Compliance Program](https://developer.microsoft.com/microsoft-365/app-compliance-program) and [Microsoft 365 App Compliance Program overview](/microsoft-365-app-certification/overview). 

The following are the basic steps:

1. [Join the certification program (recommended)](#join-the-certification-program-recommended)
1. [Update your installation executable (required)](#update-your-installation-executable-required)

#### Join the certification program (recommended)

We recommend that you join the [developer certification program](/microsoft-365-app-certification/docs/certification). Among other things, this enables your installation program to run smoother. For more information, see the following articles:

- [Get Started in Partner Center for Microsoft 365, Teams, SaaS, and SharePoint apps](/microsoft-365-app-certification/docs/userguide)
- [Microsoft 365 App Compliance Program](https://developer.microsoft.com/microsoft-365/app-compliance-program)
- [Microsoft 365 App Compliance Program overview](/microsoft-365-app-certification/overview)
- [Microsoft 365 Certification sample evidence guide overview](/microsoft-365-app-certification/docs/seg2_overview)

#### Update your installation executable (required)

The following are the steps for updating your installation executable.

1. [Check that user's Office version supports the add-ins (recommended)](#check-that-users-office-version-supports-web-add-ins-recommended)
1. [Check for AppSource disablement (recommended)](#check-for-appsource-disablement-recommended)
1. [Create a registry key for the add-in (required)](#create-a-registry-key-for-the-add-in-required)
1. [Include privacy terms in your terms & conditions (required for certified developers)](#include-privacy-terms-in-your-terms--conditions-required-for-certified-developers)

##### Check that user's Office version supports web add-ins (recommended)

We recommend that your installation check whether the user has the Office application (Excel, PowerPoint, or Word) installed and whether the Office application is a build that supports including a web add-in in a Windows application installation. If it is an old version that doesn't support web add-ins, the installation program should skip all the remaining steps. Consider displaying a message to the user that recommends that they install or update to the latest version of Microsoft 365 so they can take advantage of your web add-in. They would need to rerun the installation after installing or upgrading. 

The exact code needed depends on the installation framework and the programming language that you are using. The following is an example of how to check using C#. 

> [!NOTE]
> The installation can be designed to install the add-in for all users of the computer, if an administrator of the computer is running the installation program. To implement that design, update the code to do the following.
>
> - Check if the user is an administrator of the computer.
> - If the user is an administrator, the code should do one of the following.
>
>     - If you want to force the add-in to be installed for all users, the code should set the `supportLocalComputer` variable to `true`.
>     - If you want to give the administrator a choice between installing the add-in only for his or her self or for all users on the computer, the code should present a dialog to the administrator, return the administrator's choice, and set the `supportLocalComputer` variable accordingly.

```csharp
using Microsoft.Win32;
using System;

namespace SampleProject
{
    internal class IsBuildSupportedSample
    {
        /// <summary>
        /// This function checks if the build of the Office application supports web add-ins. 
        /// </summary>
        /// <returns> Returns true if the supported build is installed, and false if an old, unsupported build is installed or if the app is not installed at all.</returns>
        private bool IsBuildSupported()
        {
            RegistryKey hklm = Registry.CurrentUser;
            string basePath = @"Software\Microsoft\Office";
            RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(basePath);
            string wxpName = "Word"; // Can be one of "Word", "Powerpoint", or "Excel".
            bool supportLocalComputer = false; // true means LOCAL_MACHINE support, false CURRENT_USER support


            const string buildNumberStr = "BuildNumber"; 
            const int smallBuildNumber = 18227; // This is the minimum build that supports installation of a web add-in in the installation of a Windows app.
            if (supportLocalComputer)
            {
              smallBuildNumber = 18730; // This is the minimum build that supports installation of a web add-in, for all users of the computer, in the installation of a Windows app.
            }
            const int supportedBuildMajorNumber = 16; // 16 is the lowest major build of Office applications that supports web add-ins.

            if (baseKey != null)
            {
                Version maxVersion = new Version(supportedBuildMajorNumber, 0); // Initial value for the max supported build version
                foreach (string subKeyName in baseKey.GetSubKeyNames())
                {
                    if (Version.TryParse(subKeyName, out Version version))
                    {
                        if (version > maxVersion)
                        {
                            maxVersion = version;
                        }
                    }
                }

                string maxVersionString = maxVersion.ToString();
                // The Office application's build number is under this path.
                RegistryKey buildNumberKey = hklm.OpenSubKey(String.Format(@"Software\Microsoft\\Office\{0}\\Common\Experiment\{1}", maxVersionString, wxpName));

                if (maxVersion.Major >= supportedBuildMajorNumber && buildNumberKey != null)
                {
                    object buildNumberValue = buildNumberKey.GetValue(buildNumberStr);
                    if (buildNumberValue != null && Version.TryParse(buildNumberValue.ToString(), out Version version))
                    {
                        if (version.Major > supportedBuildMajorNumber || (version.Major == supportedBuildMajorNumber && version.Build >= smallBuildNumber))
                        {
                            // Build is supported
                            return true;
                        }
                        else
                        {
                            // Office is installed, but the build is not supported.
                            return false;
                        }
                    }
                    else
                    {
                        // There is no build number, which is an abnormal case.
                        return false;
                    }
                }
                else
                {
                    // An old version is installed.
                    return false;
                }
            }
            else
            {
                // Office is not installed.
                return false;
            }
        }
    }
}
```

##### Check for AppSource disablement (recommended)

We recommend that your installation check whether the [AppSource](https://appsource.microsoft.com/) store is disabled in the user's Office application. Microsoft 365 Administrators sometimes disable the store. If the store is disabled, the installation program should skip all the remaining steps. Consider displaying a message to the user that recommends that they contact their administrator about your web add-in. They would need to rerun the installation after the store is enabled. 

The following is an example of how to check for disablement of the store. 

```csharp
using Microsoft.Win32;
using System;

namespace SampleProject
{
    internal class IsStoreEnabledSample
    {
        /// <summary>
        /// This function checks if the store is enabled.
        /// </summary>
        /// <returns> Returns true if it store is enabled, false if store is disabled.</returns>
        private bool IsStoreEnabled()
        {
            RegistryKey hklm = Registry.CurrentUser;
            string basePath = @"Software\Microsoft\Office";
            RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(basePath);
            const int supportedBuildMajorNumber = 16;

            if (baseKey != null)
            {
                Version maxVersion = new Version(supportedBuildMajorNumber, 0); // Initial value for the maximum supported build version.
                foreach (string subKeyName in baseKey.GetSubKeyNames())
                {
                    if (Version.TryParse(subKeyName, out Version version))
                    {
                        if (version > maxVersion)
                        {
                            maxVersion = version;
                        }
                    }
                }

                string maxVersionString = maxVersion.ToString();

                // The StoreDisabled value is under this registry path.
                string antoInstallPath = String.Format(@"Software\Microsoft\Office\{0}\Wef\AutoInstallAddins", maxVersionString);
                RegistryKey autoInstallPathKey = Registry.CurrentUser.OpenSubKey(autoInstallPath);

                if (autoInstallPathKey != null)
                {
                    object storedisableValue = autoInstallPathKey.GetValue("StoreDisabled");

                    if (storedisableValue != null)
                    {
                        int value = (int)storedisableValue;
                        if (value == 1)
                        {
                            // Store is disabled
                            return false;
                        }
                        else
                        {
                            // Store is enabled
                            return true;
                        }
                    }
                    else
                    {
                        // No such key exists since the build does not have the value, so the store is enabled.
                        return true;
                    }
                }
                else
                {
                    // The registry path does not exist, so the store is enabled.
                    return true;
                }
            }
            else
            {
                // Office is not installed at all.
                return false;
            }
        }
    }
}
```

##### Create a registry key for the add-in (required)

Include in the installation program a function to add *one* of the following keys and values to the Windows Registry, depending on whether the add-in is being installed for all users of the computer or only for the user that is running the installation program.

```
// Only the current user
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\AutoInstallAddins\{{OfficeApplication}}\{{add-inName}}] 
"AssetIds"="{{assetId}}"

// All users of the computer
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\AutoInstallAddins\{{OfficeApplication}}\{{add-inName}}] 
"AssetIds"="{{assetId}}"
```

Replace the placeholders as follows:

- `{{OfficeApplication}}` with the name of the Office application that the add-in should be installed in. Only `Word`, `Excel`, and `PowerPoint` are supported.

   > [!NOTE]
   > If the add-in's manifest is configured to support more than one Office application, replace `{{OfficeApplication}}` with any *one* of the supported applications. Don't create separate registry entries for each supported application. The add-in will be installed for all the Office applications that it supports. 

- `{{add-inName}}` with the name of the add-in; for example `ContosoAdd-in`.
- `{{assetId}}` with the AppSource asset ID of your add-in, such as `WA999999999`.

The following is an example.

```
// Only the current user
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\AutoInstallAddins\Word\ContosoAdd-in] 
"AssetIds"="WA999999999"

// All users of the computer
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\AutoInstallAddins\Word\ContosoAdd-in] 
"AssetIds"="WA999999999"
```

The exact code will depend on your installation framework and programming language. The following is an example in C#.

> [!NOTE]
> To install the add-in for all users, change this code so that `WriteRegisterKeys` takes a bool parameter. The method should set the `supportLocalMachine` variable to the value that is passed: `true` to install for all users, `false` to install for only the current user.

```csharp
using Microsoft.Win32;
using System;

namespace SampleProject
{
   internal class WriteRegisterKeysSample
   {
        /// <summary>
        /// This function writes information to the registry that will tell Office applications to install the web add-in.
        /// </summary>

        private void WriteRegisterKeys()
        {
            bool supportLocalMachine = false; // false = CurrentUser, true = LocalMachine
            RegistryKey targetRootKey = supportLocalMachine ? Registry.LocalMachine : Registry.CurrentUser;

            string basePath = @"Software\Microsoft\Office";
            using (RegistryKey baseKey = targetRootKey.OpenSubKey(basePath))
            {
                if (baseKey == null)
                {
                    Console.WriteLine("Base registry path not found.");
                    return;
                }

                string wxpName = "Word";  // Can be "Word", "Powerpoint", or "Excel"
                string assetID = "WA999999999"; // AppSource asset ID of your web add-in
                string appName = "ContosoAddin"; // Your web add-in name
                const int supportedBuildMajorNumber = 16;
                const string assetIdStr = "AssetIDs";

                Version maxVersion = new Version(supportedBuildMajorNumber, 0);
                foreach (string subKeyName in baseKey.GetSubKeyNames())
                {
                    if (Version.TryParse(subKeyName, out Version version))
                    {
                        if (version > maxVersion)
                        {
                            maxVersion = version;
                        }
                    }
                }

                string maxVersionString = maxVersion.ToString();
                string subKeyPath = supportLocalMachine
                ? $@"Software\Microsoft\Office\{maxVersionString}\AutoInstallAddins\{wxpName}\{appName}"
                : $@"Software\Microsoft\Office\{maxVersionString}\Wef\AutoInstallAddins\{wxpName}\{appName}";

                using (RegistryKey addInKey = targetRootKey.CreateSubKey(subKeyPath))
                {
                    if (addInKey != null)
                    {
                        addInKey.SetValue(assetIdStr, assetID);
                    }
                    else
                    {
                        Console.WriteLine("Failed to create or open the registry subkey.");
                    }
                }
            }
        }

    }
}
```

##### Include privacy terms in your terms & conditions (required for certified developers)

Skip this section if you aren't a member of the certification program, but *it is required if you are*.

Include in the installation program code to add *one* of the following keys and values to the Windows Registry, depending on whether the add-in is being installed for all users of the computer or only for the user that is running the installation program.

```
// Only the current user
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\AutoInstallAddins\{{OfficeApplication}}\{{add-inName}}] 
"HasPrivacyLink"="1"

// All users of the computer
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\AutoInstallAddins\{{OfficeApplication}}\{{add-inName}}] 
"HasPrivacyLink"="1"
```

Replace the `{{OfficeApplication}}` and `{{add-inName}}` placeholders exactly as in the preceding section. The following is an example.

```
// Only the current user
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\AutoInstallAddins\Word\ContosoAdd-in] 
"HasPrivacyLink"="1"

// All users of the computer
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\AutoInstallAddins\Word\ContosoAdd-in] 
"HasPrivacyLink"="1"
```

To implement this, just make two small changes in the code sample in the previous section. 

1. To the list of `const`s at the top of the `WriteRegistryKeys` method, add the following line:

   ```csharp
   const string hasPrivacyLinkStr = "HasPrivacyLink"; // Indicates that your installer has a privacy link.
   ```

1. Just below the line `AddInNameKey.SetValue(assetIdStr, assetID);`, add the following lines:

   ```csharp
   // Set this value if the Privacy Consent has been shown on the main app installation program, this is required for a silent installation of the web add-in.
   AddInNameKey.SetValue(hasPrivacyLinkStr, 1);
   ```

#### The user's installation experience

When an end user runs your installation executable, their experience with the web add-in installation depends on two factors.

- Whether you're a [certified Microsoft 365 developer](/microsoft-365-app-certification/docs/certification).
- The security settings made by the user's Microsoft 365 administrator.

If you're certified and the administrator has enabled automatic approval for all apps from certified developers, then the web add-in is installed without the need for any special action by the user after the installation executable is started. If you're not certified or the administrator hasn't granted automatic approval for all apps from certified developers, then the user is prompted to approve inclusion of the web add-in as part of the overall installation. After installation, the web add-in is available to the user in Office on the web as well as Office on Windows.

If you're combining the installation of a web add-in with a COM/VSTO add-in, you need to think about the relationship between the two. For more information, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

## Related content

- [Make your solutions available in Microsoft AppSource and within Office](/partner-center/marketplace/submit-to-appsource-via-partner-center)
- [What is Microsoft AppSource?](/marketplace/appsource-overview)
