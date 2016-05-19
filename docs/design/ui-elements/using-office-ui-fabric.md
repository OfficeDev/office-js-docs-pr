
#Use Office UI Fabric in Office Add-ins

If you are building an Office Add-in, we encourage you to use [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) to create your user experience. The following steps walk you through the basics for using Fabric.  

##1. Set up Fabric
Add the following lines to your HTML in the head section to reference Fabric from the CDN.

     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
	 <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">


##2. Use Fabric icons and fonts
Using icons is simple. All you have to do is use an "i" element and reference the appropriate classes. You can control the size of the icon by changing the font size.

    <i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>


##3. Use styles for simple components
Fabric comes with styles for various UI elements, such as buttons and check boxes. All you have to do is reference the appropriate classes to add the corresponding style, as shown in the following example.

    <button class="ms-Button" id="get-data-from-selection">
    <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
    <span class="ms-Button-label">Get Data from selection</span>
    <span class="ms-Button-description">Get Data from the document selection</span>
    </button>

##4. Use components with sample behavior
Fabric includes some components that support behaviors (such as what happens on click). To get you started, Fabric includes some **sample code** in the form of JQuery UI plug-ins that you can use. You can also use any other framework you want to wire things up. If you do opt to use the samples, note that the code is not distributed as part of the CDN, so you have to download it from the latest release of the [Fabric GitHub project](https://github.com/OfficeDev/Office-UI-Fabric/releases), reference it, and then initialize it in your code. 

For example, to use the SearchBox component:

1. Download the SearchBox component from [GitHub](https://github.com/OfficeDev/Office-UI-Fabric/tree/master/src/components/SearchBox).
2. Add the following reference to your code: `<script src="SearchBox/Jquery.SearchBox.js"></script>`
3. Initialize the component by making sure this line executes when your page is loaded: `$(".ms-SearchBox").SearchBox();`. We recommend that you include this in the `Office.Initialize` block of your add-in.     

**Note:** If you don't intend to use all the Fabric components, you can reduce the size of the resources you download by opting instead to host the individual CSS files for each component. You can get the CSS files from the component folders in the [Fabric GitHub repository](https://github.com/OfficeDev/Office-UI-Fabric). 


##Next steps
If you're looking for end-to-end samples that show you how to use Fabric, we've got you covered. See the [Office Add-in Fabric UI sample](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample). You can also explore the interactive [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) website.

