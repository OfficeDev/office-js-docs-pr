
# Use the Exchange token validation library

You can identify the clients of your Outlook add-in by using an identity token that your add-in requests from a server running Exchange Server 2013. The token, formatted as a JSON Web token, provides a unique identifier for an email account on an Exchange server. The Exchange Web Services (EWS) Managed API provides helper classes to simplify the use of the identity token.

## Prerequisites for using the validation library


To validate an Exchange identity token, you must have the EWS Managed API authentication library and the Windows Identity Foundation (WIF), along with a DLL that extends the WIF with handlers for JSON tokens. Make sure that you download the following resources:


- [Exchange Web Services Managed API](http://go.microsoft.com/fwlink/?LinkID=255472)
    
- [Windows Identity Foundation ](http://www.microsoft.com/en-us/download/details.aspx?id=17331)
    
- [Windows.IdentityModel.Extensions.dll for 32-bit applications](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll for 64-bit applications](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## Validate the Exchange identity token


The EWS Managed API validation library provides the  **AppIdentityToken** class to manage the Exchange identity tokens. The following method shows how to create an **AppIdentityToken** instance and call the **Validate** method to verify that the token is valid.


```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

        private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
        {
            try
            {
                AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
                token.Validate(new Uri(hostUri));

                return token;
            }
            catch (TokenValidationException ex)
            {
                throw new ApplicationException("A client identity token validation error occurred.", ex);
            }
        }

```


## Additional resources



- [Authenticate an Outlook add-in by using Exchange identity tokens](../outlook/authentication.md)
    
- [Inside the Exchange identity token](../outlook/inside-the-identity-token.md)
    
- [Validate an Exchange identity token](../outlook/validate-an-identity-token.md)
    
