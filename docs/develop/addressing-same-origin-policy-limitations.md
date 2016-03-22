
# Addressing same-origin policy limitations in Office Add-ins


The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.

Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.

To overcome same-origin policy enforcement when you develop add-ins, you can:

- Use JSON/P for anonymous access. 
    
- Implement server-side script using a token-based authentication scheme.
    
- Using cross-origin resource sharing (CORS).
    
- Build your own proxy using IFRAME and POST MESSAGE.
    

## Using JSON/P for anonymous access


One way to overcome this limitation is to use JSON/P to provide a proxy for the web service. You do this by including a  **script** tag with a **src** attribute that points to some script hosted on any domain. You can programmatically create the **script** tags, dynamically create the URL to point the **src** attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.

The following is an example of JSON/P that uses a technique that will work in any Office Add-in.




```
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## Implementing server-side script using a token-based authentication scheme


Another way to address same-origin policy limitations is to implement the add-in's webpage as an ASP page that uses OAuth or caches credentials in cookies.

For an example that uses OAuth for authentication, see [Twitter SharePoint web part with OAuth](http://aidangarnish.net/post/Twitter-SharePoint-Web-Part-With-OAuth.aspx).

For an example of server-side code that shows how to use the  **Cookie** object in **System.Net** to get and set cookie values, see the [Value](http://msdn2.microsoft.com/EN-US/library/4f772twc) property.


## Using cross-origin resource sharing (CORS)


For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).


## Building your own proxy using IFRAME and POST MESSAGE


For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).


## Additional resources



- [Privacy and security for Office Add-ins](../../docs/develop/privacy-and-security.md)
    
