# Word add-ins

Welcome to the Word add-in JavaScript API documentation. The Word JavaScript API is a part of the Office add-in programming model for extending Microsoft Office applications. The add-in programming model uses web applications to host your extension to Word. You can now extend Word with any web platform or language that you prefer. 

## Get started now

Are you the type that wants to read fewer words and just wants to see the code? Then let's go and [build your first Word add-in](build-your-first-word-add-in.md). 

## API Overview

Before we start, you need to know that this new Word add-in model is different than what was available with Word in Office 2013. The previous object model was not typed and provided a generic API for extending Office clients. While this model is still applicable to Word 2016, we strongly suggest that you start using the new Word object model. This new object model provides access to familiar Word objects like: Body, Sections, Paragraphs, Fonts, Content Controls, and Ranges.



The new Word 
(link into the programming guide and reference)

 Creating and distributing a Word extension Create your extension and host it. 


, post it to a network share, a SharePoint App Catalog, or Here you'll learn how you can extend Word with 

The new JavaScript APIs for Word and Excel 2016 change the way that JavaScript interacts with objects like paragraphs, pages, ranges, worksheets, and charts that are running inside the Office application.  Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the new APIs provide “proxy” JavaScript objects that correspond to the real objects running in a separate process (or across the network in the case of Office Online).  You can directly interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.  These interactions with proxy objects aren’t immediately realized in the running script, though, so we provide a method on the context called “executeAsync()” that synchronizes the state between your running JavaScript and the real objects in Office by executing instructions queued in your script and retrieving properties of loaded Office objects for use in your script.  



Let's point you at how to get started... a simple step through that shows how to create your first addin in the simplest form. 



## Check out the code


(link to Robs code, and a code samples page). There should be a link to the store if it is in there. 

## Give feedback on the API

The documentation for this API is hosted on GitHub with the intention that we can improve the documentation by making it open for opening [issues](https://github.com/OfficeDev/office-js-docs/issues) against the documentation. Issues can include errors in the documentation, requests for clarification, or requests for improvements in the documentation. We also welcome general feedback about the API and the experience you have with it.  

## Additional links

(links into other important documentation, linke links into manifest documentation, officejS reference, 