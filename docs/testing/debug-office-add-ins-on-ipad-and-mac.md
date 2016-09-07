
# Debug Office Add-ins on iPad and Mac

You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac. 

## Debugging with Vorlon.js 

Vorlon.js is a debugger for web pages, similar to the F12 tools, that is designed to work remotely and allows you to debug web pages across different devices. For more information, see the [Vorlon website](http://www.vorlonjs.com).  

To install and set up up Vorlon: 

1.	Install [Node.js](https://nodejs.org) and [Git](https://git-scm.com/) if you haven't already. 

2.	Install Vorlon using git with the following command: `git clone https://github.com/MicrosoftDX/Vorlonjs.git`.

3.	Install dependencies with `npm install`.

4.	Add-ins require HTTPS, so by extension any scripts that they use must be HTTPS as well, including the Vorlon script. Therefore, you have to configure Vorlon to use SSL in order to use Vorlon with add-ins. Under the folder where you installed Vorlon, go to the /Server folder and edit the config.json file. Change the **useSSL** property to **true**. While you're there, you can also enable the plugin for Office Add-ins (change its "enabled" property to true). 

5.	Run the Vorlon server with the command `sudo vorlon`. 

6.	Open a browser window and go to [http://localhost:1337](http://localhost:1337), which is the Vorlon interface. Trust the security certificate, which you should be prompted to do. You can also find the security certificate in the Vorlon folder under /Server/cert. 

7.	Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

Now, whenever you open the add-in on a device, it will show up in the list of Clients in Vorlon (on the left side of the Vorlon interface). You can remotely highlight DOM elements, remotely execute commands, and much more.  

![Screenshot that shows the Vorlon.js interface](../../images/vorlon_interface.png)

The Office plugin adds extra capabilities for Office.js, such as exploring the object model and executing Office.js calls. 
