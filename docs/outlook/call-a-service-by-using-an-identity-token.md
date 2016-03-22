
# Call a service from an Outlook add-in by using an identity token in Exchange

An identity token provides a unique identifier for each of your customers that you can use to personalize the service that you provide. Your code can ask the Exchange server for an identity token by using an asynchronous method call that returns a string to your Outlook add-in. The string contains a JSON Web Token (JWT) identity token. Your add-in doesn't need to unpack the token. Instead, it passes the token on to your web service so that your service can authenticate the request from the add-in.

The web service that supports your add-in must run on the same server that hosts the add-in HTML and JavaScript source files. This prevents cross-site scripting errors. Your server can proxy the request on to other web services if your application requires it.

Adding an identity token to the service request that your add-in sends is easy. You request the token, use the token, and then use the web service response. Here's how it looks with a simple XML document that you send to your server by using the  **XmlHttpRequest** method.

## Request a token from your Exchange server


This simple initialization method for an add-in uses the  **getUserIdentityTokenAsync** method to request an identity token from the Exchange server. The _getUserIdentityToken_ parameter is the function that is called when the asynchronous request to the server returns. See the next step for the callback method.


```js
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## Use the identity token


The callback function for the  **getUserIdentityTokenAsync** method has one parameter that contains the user identity token in its **value** property.

This callback function creates an  **XMLHttpRequest** object to call the web service. Set the **onreadystatechange** property on the **XMLHttpRequest** object to the name of the function that should run when your add-in gets a response from the web service.




```js
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## Use the web service response


This is another simple function that processes the response from the web service. It follows the standard pattern for  **XHMHttpResponse** callback functions. It waits for the entire response to come in from the web service and then puts the contents of the response on the add-in UI. The response that this function is parsing is the response from the web service. For information about this response, see [Validate an Exchange identity token](../outlook/validate-an-identity-token.md). 


```js
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## Example: Calling a web service with identity tokens


Identity tokens provide identity information about the client that is calling your service to a web service that is running on your server. To use identity tokens, you'll need the following:


- An Outlook add-in that requests an identity token from the Exchange server and sends in on to your web service. The information in this topic will help you create that add-in.
    
- A web service running on the server that provides the UI for your add-in that validates the identity token. You'll find the information that you need to create the web service in one of the following topics:
    
      - [Use the Exchange token validation library](../outlook/use-the-token-validation-library.md) -- If you're using the validation library that we provide.
    
  - [Validate an Exchange identity token](../outlook/validate-an-identity-token.md) -- If you're writing your own validation code.
    

### Code for the sample add-in


The following files are required for the add-in described in this article:


- IdentityTest.js - The JavaScript files that provide the business logic for the add-in.
    
- IdentityTest.html - The HTML file that provides the UI for the add-in.
    
You'll also need the Identity Test web service. For information about that web service, see [Validate an Exchange identity token](../outlook/validate-an-identity-token.md).


#### IdentityTest.js

The following example shows the IdentityTest.js file.


```js
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### IdentityTest.html

The following example shows the IdentityTest.html file.


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## Next steps


Now that you know how to request an identity token, you need to use the token on the server side of the request. The following articles will help you get started:


- [Use the Exchange token validation library](../outlook/use-the-token-validation-library.md)
    
- [Validate an Exchange identity token](../outlook/validate-an-identity-token.md)
    
- [Authenticate a user with an identity token for Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## Additional resources



- [Authenticate an Outlook add-in by using Exchange identity tokens](../outlook/authentication.md)
    
- [Inside the Exchange identity token](../outlook/inside-the-identity-token.md)
    
