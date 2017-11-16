# auth.getAccessTokenAsync method
Requests an authentication Token for the user currently signed into Office.

> **Important:** This API currently works only in Excel, Outlook, PowerPoint, Word, and OneNote in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview).

## Requirements

This method is available in the IdentityAPI [requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md). To specify the IdentityAPI requirement set, use the following in your manifest.

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="IdentityAPI"/> 
   </Sets> 
 </Requirements> 

```

To detect this API at runtime, use the following code.

```js
 if (Office.context.requirements.isSetSupported('IdentityAPI', 1.1)) 
 	{  
    	 // Request an SSO Token 
 	} 
 else 
	 { 
	     // Alternate path 
	 } 
```

## Syntax

```js
getAccessTokenAsync([Options,] callback);
```

## Examples

```js
	Office.context.auth.getAccessTokenAsync(function(result) {
	    if (result.status === "succeeded") {
		var token = result.value.accessToken;
		...
	    } else {
		console.log("Error obtaining token", result.error);
	    }
	});
```

## Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|Options|object|Optional. An options object, described below|
|callback|string|Accepts a callback method to handle the identity objects.|

### Options object
The following actions are available during authentication.

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|**authChallenge**|string|Optional. Prompts the user to sign in with additional authentication requirements.|
|**forceConsent**|bool|Optional. Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has been revoked.|
|**forceAddAccount**|bool|Optional. Prompts the user to add (or to switch if already added) his or her Office account.|
|**asyncContext**|any|Optional. A user-defined item of any type that is returned in the AsyncResult object without being altered.|

### callback method
When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](/reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getAccessTokenAsync** method, you can use the properties of the **AsyncResult** object to access the access token.

## Remarks

This API requires a single sign-on configuration that bridges the add-in to an Azure application. Office users sign-in with Organizational Accounts and Microsoft Accounts. Microsoft Azure returns tokens intended for both user account types to access resources in the Microsoft Graph.

## Support details

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).

