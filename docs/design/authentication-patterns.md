# Authentication
Use  dialogs to present authentication screens to your end users. Our flows provide a seamless experience that integrates the Microsoft sign-in with your brand. 

## Best practices

|Do                                                  |Don't                   |
|:-----------|:-----------|
|Utilize the space to reinforce your company's brand. Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your services. [Learn More](https://azure.microsoft.com/en-us/documentation/articles/active-directory-add-company-branding/).|Don't overcrowd the task pane or dialogs with buttons.|
| Target your dialogs to a few key widths or breakpoints for better responsive design. [Learn more](https://msdn.microsoft.com/windows/uwp/layout/screen-sizes-and-breakpoints-for-responsive-design)|                         

Apply the following patterns as applicable to create or enhance the authentication experience for your add-in. 

## Authentication flow for Single Sign On (SSO)


While Single Sign On (SSO) support has not yet officially been released, once published it should be considered for the default authentication flow of your add-in.  SSO provides the user with the best authentication experience, esentially moving the authentication into the broader installation process.

As an add-in is being installed, a user will see a consent window similar to the one below:

![Authentication Flow - Single Sign On](../Screens/Components/Single_Sign_On_Consent@2x.png)

The add-in publisher will have control over the logo and strings included in the SSO window, but the UI is pre-configured by Microsoft.

After consent has been given by the user, the installation process will conclude, and your add-in will be added and ready to use.

![Authentication Flow - Single Sign On](../Screens/Addin_Screens/SSO_Modal@2x.png)

![Authentication Flow - Single Sign On](../Screens/Addin_Screens/TaskPane_Opened@2x.png)

## Authentication flow for a single identity provider

If SSO is not available to a user, an alternative authentication flow is for a single or multi-identity provider.
=======

![Authentication Dialog Single Identity - Flowchart](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_single_flow.png)

1. First Run Placemat - Place the branded sign-in button as a clear call-to action inside your add-in's UI.
![Authentication Flow - First run placemat](../Screens/Addin_Screens/FRE-Value@2x.png)

2. Provider Sign-in - The identity provider will have their own UI. Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service. [Learn More](https://azure.microsoft.com/en-us/documentation/articles/active-directory-add-company-branding/).

![Authentication Dialog Single Identity - Provider Sign-in](../Screens/Addin_Screens/Multi_Authentication_Modal@2x.png)

3. Progress - Indicate progress while settings and UI load.
![Authentication Dialog Single Identity - Progress](../Screens/Addin_Screens/Multi_Authentication_Modal_Interstitial@2x.png)
=======

## More details

- **Microsoft branded sign-in button** - When using Microsoft's Identity service you'll have the opportunity to use a branded sign-in button that is customizable to light and dark themes. [Learn more](https://azure.microsoft.com/en-us/documentation/articles/active-directory-branding-guidelines/#visual-guidance-for-sign-in).
