
# Authenticate an Outlook add-in by using Exchange identity tokens

Your Outlook add-in can provide your customers with information from anywhere on the Internet, whether from the server that hosts the add-in, from your internal network, or from somewhere else in the cloud. If that information is protected, however, your add-in needs a way to associate the Exchange email account with your information service. Exchange 2013 can enable single sign-on (SSO) for your add-in by providing a token that identifies the email account that is making the request. You can associate this token with a registered user for your application so that the user is recognized whenever the add-in connects to your service.

## Identity tokens


Two of our sample add-ins use publically available information - one shows a Bing map for addresses in a message, and one shows a preview for YouTube video links in a message. But your add-in can also access nonpublic information. You can use the server that hosts your add-in to link your add-in to the information in your internal network, or anywhere in the cloud.

You can use many different techniques to identify and authenticate add-in users. Exchange 2013 simplifies user authentication by providing your add-in an identity token that identifies a specific Exchange email account. You can associate this token in your service with a registered user, enabling single sign-on (SSO) for your customers that use Outlook add-ins. 

To use SSO in your add-in, the code does this:


* Calls a function in the Outlook add-in API that returns an identity token.
* Sends the token together with a request to your server.
* Unpacks the response from the server to display information from your service.
    
On the server side, things are somewhat more complex. When your server receives a request from an Outlook add-in, the process works like this:

* The server validates the token. You can use our [managed token validation library](../../docs/outlook/use-the-token-validation-library.md), or you can [create your own library](../../docs/outlook/validate-an-identity-token.md) for your service.
* The server looks up the unique identifier from the token to see whether it's associated with a known identity. Your service must [implement a method that matches the identifier](../../docs/outlook/authenticate-a-user-with-an-identity-token.md) with known users of your service.
* If the unique identifier matches an identifier previously stored with a set of credentials on the server, your server can respond with the requested information without requiring your customer to log on to your service.
* If the unique identifier is unknown, the server sends a response asking the user to log on with credentials for the server.
* If the credentials match a known identity on the server, you can map that identity to the unique identifier in the token so that the next time a request comes in, your server can respond without requiring an additional logon step.

 >**Note**  This is just one suggestion for how to use the identity token. As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.

Let's get into the specifics. In the following articles, we'll use a simple Outlook add-in that sends the identity token and a list of phone numbers found in the message to a web service. 

- [Inside the Exchange identity token](../outlook/inside-the-identity-token.md)
- [Call a service from an Outlook add-in by using an identity token in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
- [Use the Exchange token validation library](../outlvalidate-an-identity-token.md ook/use-the-token-validation-library.md)
- [Validate an Exchange identity token](../outlook/validate-an-identity-token.md )
- [Authenticate a user with an identity token for Exchange](../outlook/validate-an-identity-token.md)


## Additional resources



- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Call web services from an Outlook add-in](../outlook/web-services.md)
    


