
# Package your add-in using Napa or Visual Studio to prepare for publishing

Your Office Add-in package contains an XML file that you'll use to publish the add-in. You'll have to publish the web application files of your project separately.

## Package an Office Add-in that you create by using Napa



1. In Napa, on the side of the page, choose the  **Publish** button (
![Publish button](../../images/Apps_NAPA_Publish.png)).
    
2. In the  **Publish settings** dialog box, choose **Next**.
    
3. Provide the URL of the website that will host the content files of your add-in (for example, the default HTML and JavaScript files of your project), and then choose  **Publish**.
    
4. In the  **Publish successful** dialog box, choose the **Publish location** link.
    
    A document library appears that contains the XML manifest file of your add-in, and the web content files. 
    
Next, manually copy the web content files of (style sheets, JavaScript files, and HTML files) to the web server that hosts the website that you provided in the  **Publish settings** dialog box.

You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). 


## Deploy your web project and package your add-in by using Visual Studio 2015



### To deploy your web project


1. In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.
    
    The  **Publish your add-in** page appears.
    
2. In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.
    
     >**Note**  A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.

    If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.
    
    For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. In the  **Publish your add-in** page, choose the **Deploy your web project** link.
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### To package your add-in


1. In the  **Publish your add-in** page, choose the **Package the add-in** link.
    
    The  **Publish Office and SharePoint Add-ins** wizard appears.
    
2. In the  **Where is your website hosted?** dropdown list, select or enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.
    
    You have to specify an address that begins with the HTTPS prefix to complete this wizard. In general, using an HTTPS endpoint for your website is the best approach, but it is not required if you don't plan to publish your add-in to the Office Store. After the package is created, you can open the manifest in Notepad and replace the HTTPS prefix of your website with an HTTP prefix. For more information, see [Why do my add-ins have to be SSL-secured?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**Note**  Azure websites automatically provide an HTTPS endpoint.

    Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder. 
    
If you plan to submit your add-in to the Office Store, you can choose the  **Perform a validation check** link to identify issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.

You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in  `OfficeAppManifests` in the `app.publish` folder. For example:

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## Additional resources



- [Publish your Office Add-in](../publish/publish.md)
    
- [Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
