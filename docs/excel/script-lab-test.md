---
title: Testing Script Lab integration
description: This sample test file demonstrates an upcoming ScriptLab feature that will enable developers to try out their snippets in Excel, Word, and PowerPoint.
ms.date: 03/14/2018
---


# Testing Script Lab integration

This sample test file demonstrates an upcoming ScriptLab feature that will enable developers to try out their snippets in Excel, Word, and PowerPoint. 

## Prerequisites

- You'll need a View URL from a ScriptLab snippet.

> [!NOTE] 
> We *should* indicate that ScriptLab needs Office 365 to explore the most recent snippets. Developers can obtain an Office 365 developer subscription through our [Office 365 Developer Program](https://developer.microsoft.com/en-us/office/dev-program) for development purposes only. 
> See the [Office 365 Developer Program documentation](https://docs.microsoft.com/en-us/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription. 


## Try it out button

In this way, we will add a **Try it out** button, which we recommend be associated with a code snippet. To enable this, we are using an Office UI Fabric class to style a link as a button. On the link itself, remember to set the `aria label` attribute.

### Demo

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</button>


### Code

```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## Embed script lab as an iframe

In this mode, we will embed a snippet directly as an iframe into our documents. The width has been set at 95% (based on the width of all other snippets) and we recommend that you remove the fameborder of the iframe.  Height typically should be adjusted to match the snippet.

### Demo

<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

### Code

```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## Testing considerations

We need to verify mobile, non-Office 365 subscriptions (we have feedback on office-js-docs where lots of developers were on 2013 or earlier).  

For the Embed path, we need final sign-off and need to make sure the content exposed in the view gist page meets our accessibility guidelines.


