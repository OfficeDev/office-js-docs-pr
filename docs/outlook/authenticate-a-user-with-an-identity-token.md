
# Authenticate a user with an identity token for Exchange

You can implement a single sign-on (SSO) authentication scheme for an information service that enables customers who are using Outlook add-ins to connect to your service by using their Exchange server credentials. This article shows how to match credentials by using a simple  **Dictionary** object-based user data store.

 >**Note**  This is just a simple example of SSO and should not be used in your production code. As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.


## Prerequisites for using SSO authentication


To use an identity token for SSO, your service application needs to have a valid identity token. You can learn about identity tokens, and how to request and validate an identity token, in the following articles:


- [Inside the Exchange identity token](../outlook/inside-the-identity-token.md)
    
- [Call a service from an Outlook add-in by using an identity token in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Use the Exchange token validation library](../outlook/use-the-token-validation-library.md) if you are using managed code, or [Validate an Exchange identity token](../outlook/validate-an-identity-token.md) if you are writing your own token validation method.
    

## Authenticate a user


The following code example shows a simple authentication object that matches the unique identity represented by an identity token with a set of credentials for a service. The  **TokenAuthentication** class provides a method, **GetResponseFromService**, that will return a response for previously authenticated tokens, or ask the user to provide credentials that can be authenticated and associated with the identity token. The code is not complete; it assumes that you will provide the following objects and methods.



|**Object/method**|**Description**|
|:-----|:-----|
|**LocalCredentials** object|Represents the user's credentials for your service. The structure of the object depends on the requirements of your service.|
|**IdentityToken** object|Contains a user identity token sent to your service by an Outlook add-in. The object must contain at least the unique Exchange identifier of the user and the authentication metadata URL for the server that issued the token. This example uses the identity token object defined in the article [Validate an Exchange identity token](../outlook/validate-an-identity-token.md).|
|**JsonResponse** object|Represents the response from your service. The object can be serialized to a JSON object.|
|**CallService** method|Calls your service with a  **LocalCredentials** object that contains the user's credentials for the service and an object that contains data for the service request. If the credentials are valid, this method returns a **JsonReponse** object that contains the results of the request. If the credentials are not valid, this method returns **null**.|
|**GetCredentialsResponse** method|Returns a  **JsonReponse** object that your mail Office Add-in will recognize as a request for credentials for the service.|
|**LocalCredentialsAreValid** method|Returns  **true** if the credentials supplied to the service are valid; otherwise, it returns **false**.|

 >**Note**  This is just one suggestion for how to use the identity token. As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.


```C#
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## Authenticating a user with the managed validation library


If you are using the managed library to validate identity tokens, you do not need to compute the unique key. The  **UniqueUserIdentification** property on the **AppIdentityToken** class can be used directly as the unique key for the user. The following code example shows the modifications to the **GetResponseFromService** method in the previous example that you need to make to use the **AppIdentityToken** class.


```js
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## Additional resources



- [Authenticate an Outlook add-in by using Exchange identity tokens](../outlook/authentication.md)
    
- [Call a service from an Outlook add-in by using an identity token in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Use the Exchange token validation library](../outlook/use-the-token-validation-library.md)
    
- [Validate an Exchange identity token](../outlook/validate-an-identity-token.md)
    
