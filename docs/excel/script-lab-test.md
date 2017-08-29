# Testing Script Lab Integration

This is a sample test file, meant to demonstrate an upcoming ScriptLab feature which will enable developers to try out their snippets in Excel, Word, PowerPoint.  

## Pre-reqs:
- You'll need a View URL from a ScriptLab snippet
- Note: We *should* indicated ScriptLab needs Office 365 to explore the most recent snippets.  Developers can obtain a Office 365 Subscription through our [Office 365 developer program](https://dev.office.com/devprogram), for development purposes only.  


## Try it out 'Button'
In this way, we will add a Try it out button, which we recommend is associated with a code snippet.  To enable this, we are using a Office UI Fabric class to style a link as a button. On the link itself, remember to set the *aria label* atrribute.

**Demo:**

<a href="https://dev.microsoft.com" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>


**Code:**
```html
<a href="<add link to protocal handler or 'online' scriptlab version if available" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## Embed script lab as an iframe
In this mode, we will embed a snippet directly as an iframe into our documents. The width has been set to take 100% and we recommend you remove the fameborder of the iframe.  Height typically should be adjusted to match the snippet.

**Demo:**
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="100%" frameborder="0"></iframe>

**Code:**
```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="100%" frameborder="0"></iframe>
```

## Testing Considerations
We need to verify mobile, non-Office 365 subscriptions (we have feedback on the office js docs where lots of developers were one 2013 or earlier.  

For the Embed path, we need final sign off and need to make sure the content exposed in the view gist page meets our Accessibility guidelines.
