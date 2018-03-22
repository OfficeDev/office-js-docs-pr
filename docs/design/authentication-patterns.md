# Authentication
Use  dialogs to present authentication screens to your end users. Our flows provide a seamless experience that integrates the Microsoft sign-in with your brand. 

## Best practices

|Do                                                  |Don't                   |
|:-----------|:-----------|
|Utilize the space to reinforce your company's brand. Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your services. [Learn More](https://azure.microsoft.com/en-us/documentation/articles/active-directory-add-company-branding/).|Don't overcrowd the task pane or dialogs with buttons.|
| Target your dialogs to a few key widths or breakpoints for better responsive design. [Learn more](https://msdn.microsoft.com/windows/uwp/layout/screen-sizes-and-breakpoints-for-responsive-design)|                         

Apply the following patterns as applicable to create or enhance the authentication experience for your add-in. 

## Authentication flow for Single Sign On (SSO)

While Single Sign On (SSO) support has not yet officially been released, once released it should be considered for the default authentication flow of your add-in.  SSO provides the user with the best authentication experience, esentially moving the authentication flow into the installation process.

As an add-in is being installed, a user will see a consent modal similar to the one below:

![Authentication Flow - Single Sign On](../Screens/Components/Single_Sign_On_Consent@2x.png)

After consent has been given by the user, the installation process will conclude, and your add-in will be added and ready to use.

![Authentication Flow - Single Sign On](../Screens/Addin_Screens/SSO_Modal@2x.png)


## Authentication flow for a single identity provider

If SSO is not available to a user, an alternative authentication flow is for a single identity provider.  He

![Authentication Dialog Single Identity - Flowchart](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_single_flow.png)

1. First Run Placemat - Place the branded sign-in button as a clear call-to action inside your add-in's UI.
![Authentication Flow - First run placemat](../Screens/Addin_Screens/FRE-Value@2x.png)

2. Provider Sign-in - The identity provider will have their own UI. Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service. [Learn More](https://azure.microsoft.com/en-us/documentation/articles/active-directory-add-company-branding/).
![Authentication Dialog Single Identity - Provider Sign-in](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_single_taskPaneCallouts2.png)


3. Progress - Indicate progress while settings and UI load.
![Authentication Dialog Single Identity - Progress](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_single_taskPaneCallouts3.png)

4. Home Page - Land your users on a useful home page to begin their add-in experience.
![Authentication Dialog Single Identity - Home Page](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_single_taskPaneCallouts4.png)

5. Sign-out - Include a discoverable entry point for users to manage their profile.
![Authentication Dialog Single Identity - Sign-out](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_single_taskPaneCallouts5.png)





## Authentication flow for multiple identity providers

 Consider this UX design pattern when using multiple identity providers or your add-in has limited space to display branded sign-in buttons. 

Recommended screen flow for when using multiple identity providers in your add-in.

![Authentication Dialog Multiple Identity - Flowchart](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_flow.png)


1. First Run Placemat - The screen contains a clear call to action, "Sign-in"
![Authentication Flow - First Run Placemat](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts.png)


2. Provider Choices - End users are presented with a set of identity providers to choose from, including an authentication form. Note that the add-in UI is on hold until the dialog closes.
![Authentication Dialog Multiple Identity - Provider Choices](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts2.png)


3. Provider Sign-in - The identity provider will have their own UI. Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service. [Learn More](https://azure.microsoft.com/en-us/documentation/articles/active-directory-add-company-branding/).
![Authentication Dialog Multiple Identity - Provider Sign-in](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts3.png)


4. Progress - Indicate progress while settings and UI load. 
![Authentication Dialog Multiple Identity - Progress](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts4.png)


5. Home Page - Land your users on a useful home page to begin their add-in experience.
![Authentication Dialog Multiple Identity - Home Page](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts5.png)


6. Sign-out - Include a discoverable entry point for users to manage their profile.
![Authentication Dialog Multiple Identity - Sign-out](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts6.png)

### Variants
Provider Choices Variant A - Authentication form with multiple provider sign-in buttons.
![Authentication Dialog Multiple Identity - Provider choices variant A](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts7.png)

Provider Choices Variant B - Multiple provider sign-in buttons.
![Authentication Dialog Multiple Identity - Provider choices variant B](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/auth_multi_taskPaneCallouts8.png)


## More details

- **Microsoft branded sign-in button** - When using Microsoft's Identity service you'll have the opportunity to use a branded sign-in button that is customizable to light and dark themes. [Learn more](https://azure.microsoft.com/en-us/documentation/articles/active-directory-branding-guidelines/#visual-guidance-for-sign-in).
