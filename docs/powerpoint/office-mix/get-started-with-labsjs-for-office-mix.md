
# Get started with LabsJS for Office Mix



LabsJS content exposes an API (labs.js), samples, documentation, and associated files that you can use to develop interactive Labs, integrate them into Office Mix, and then render them in Microsoft PowerPoint. These labs are, in fact, Office Add-ins that you create using HTML5 and the labs.js JavaScript library.

## LabsJS content

LabsJS provides documentation, sample labs, and the files required to create and publish your own labs for Office Mix.


**Required files**


|**File**|**Description**|
|:-----|:-----|
|labs-1.0.4.js|The LabsJS JavaScript API for the development of Office Mix Labs. This file must be included in your project to allow it to integrate with Office Mix. The file is also hosted on a content delivery network (CDN) at  <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. When you publish your app you must link to the file on the CDN.|
|labs-1.0.4.d.ts|TypeScript definition file for labs.js. This makes it possible to easily integrate your TypeScript code with labs.js. The definition file also provides a broad overview of all the components contained in labs.js. You can download TypeScript from [http://www.typescriptlang.org/](http://www.typescriptlang.org/). The definition file was built against TypeScript version 0.9.1.1.|
|History|Release history for the various versions of the labs.js library.|
|Labshost.html|A web page that allows you to view and debug your lab against Office Mix, outside the context of PowerPoint. To use the page, type in your URL to the main input box and it will load within the frame. Data exchanged between the API and Office Mix when running in PowerPoint or the Office Mix lesson player will show up in the input boxes to the right. The data can also be pre-seeded. Note that the sample Labs in the Samples section show existing Office Mix Add-ins running in the host context.|
|SampleManifest.xml|A sample Office Add-ins manifest to use as a template for creating your own application manifest.|
|Simplelab.html|A sample Lab created with labs.js. Allows for the selection of a web page and insertion of a web page, and which then tracks the user viewing it.|
|Simplelab.ts|The TypeScript file used to create simplelab sample.|
|Simplelab.js|JavaScript version of the Simplelab sample. Both this and the simplelab.ts show use of the LabsJS API.|

## Set up your development environment

The labs.js library serves as an abstraction layer on top of the office.js library (the API for Office Add-ins), so the labs you create using the labs.js library are actually Office Add-ins. In order to work with the labs.js library and to run these labs inside Office Mix, you must first set yourself up as an Office Add-ins developer.


### Register for an Office 365 Developer Site

Your first step is to sign up for an Office 365 Developer Site. This allows you to host and test your lab before submitting it to the Office Store. The site allows you to publish your add-in to Office Mix and test it in a live environment.

For more information, see [Set up a development environment for SharePoint Add-ins on Office 365](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx). You only need to follow the first two steps; installing the "Napa" developer tools is optional.


### Set up an app catalog on SharePoint Online

After your developer site is created and provisioned, you then set up an add-in catalog on SharePoint Online. For more information, see [Set up an add-in catalog on Office 365](../../publish/set-up-an-add-in-catalog-on-office-365.md).

For Office Mix, you use an add-in catalog so you can insert pre-production add-ins into a lesson and conduct end-to-end testing before submitting the labs to the store.


## Create your lab

To create your first lab, follow the steps in the [walkthrough](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md), which explains how to create a simple true/false quiz. See [Walkthrough: Creating your first lab for Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md).


## Publish your lab

After you create your lab, you can publish it and submit it to the store.


### Create and upload your application manifest

The application manifest is an XML document that describes your LabJS lab. It provides a reference to the URL where the lab is hosted and provides details about the lab, including display name, description, icons, size, and so on.

We include a sample manifest, "SampleManifest.xml". For more information about the manifest schema as well as a link to the schema definition, see [Office Add-ins XML manifest](../../../docs/overview/add-in-manifests.md).

To upload your manifest to your SharePoint site, first navigate to your application catalog, which you'll typically find at the URL <code>https://\<your site\>/sites/AppCatalog</code>. Then, choose the  **New app** button and follow the steps to upload your application manifest.


### Update your PowerPoint 2013 catalog

Next, update your PowerPoint 2013 catalog. After that you can log on with your developer account.

Start by updating your PowerPoint 2013 catalog. Launch PowerPoint 2013 and navigate the menu path  **File > Options > Trust Center > Trust Center Settings > Trusted App Catalogs**. From there, add a reference to your app catalog, and choose  **OK**. PowerPoint 2013 will ask you to sign out for the changes to take place. Sign out.

Finally, log back on using the developer account. Choose your logon name in the upper right corner in PowerPoint 2013 and log on using your developer account. You can now insert your add-in.


### Insert, publish, and view your app

To insert your add-in into the catalog, choose the  **Insert** ribbon, then choose **Store** in the **Apps** section. Choose **My Organization**, and you will see the add-in in your add-in catalog. Choose the add-in, select  **Insert**, and you add-in (lab) is inserted in the PowerPoint 2013 document.

Now you can take advantage of all the available Office Mix functionality to publish your lesson with your new lab.


 >**Important**:  To view the application, you must log on to your SharePoint catalog with the same browser that you view your lesson from. SharePoint catalogs only allow access from authenticated users, so to see your application you need to log on first. 


### Submit your lab to the Office Store

To submit your lab to the Office Store, see [Publish your Office Add-in](../../publish/publish.md)


## Additional resources



- [Office Mix add-ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Office Add-ins](../../../docs/overview/office-add-ins.md)
    
- [Creating your first lab for Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    
