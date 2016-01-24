# KurveJS and App model V2 - Getting Started

As discussed in the <a href="../../README.md">readme document</a>, the app model v2 is a new model that eventually will become the standard way Microsoft applications deal with authentication and authorization moving forward. Currently, this model is in preview, therefore it presents several limitations and some issues. We suggest developers to only use this model for learning purposes at this stage.

This document will explain how to get started with Kurve JS and the app model V2.


## Step by step guide

###Step 1: Register your application in the new application registration portal.

Your application needs an ID. In order to get one, you need to go to <a href="https://apps.dev.microsoft.com/">https://apps.dev.microsoft.com/</a> and register a new application. After creating the application, click at "add platform" and choose "Web". Make sure you leve the implicit flow checkbox checked.

Important notes in this step:

1-Different than with the app model v1, here you won't specify any permissions during the app registration. The permissions are requested during runtime. More details below.

2-Make sure you have enabled the implicit flow as mentioned above, otherwise Kurve.JS won't work

3-Make sure the redirect URL points to where the login page will be. Below you will see as part of the guidance that you need to copy the sample login.html to your website, so if it points to something like https://localhost:8000/login.html, this will be exactly what the reply URL will also have to be.

###Step 2: Reference the JavaScript file from your HTML page:

```html
    <script src="Kurve.min-0.3.6.js"></script>
```

Alternatively you can just reference the script from our CDN:

```html
    <script src="https://kurvejs.blob.core.windows.net/dist/kurve.min-0.3.6.js"></script>
```

###Step 3: Code:


The identity and graph classes can work independently or together. For example, you may have authenticated and retrieved an access token from somewhere else and only give it to the graph library, or you can use the identity library for authentication.

```javascript

//Option 1: We already have an access token obtained from another library:
var graph = new Kurve.Graph({ defaultAccessToken: token });


//Option 2: Automatically linking to the Identity object (note we specify we are working with the app model V2)
var identity = new Kurve.Identity(
        "your_client_id",
        "your_redirect_url",
        Kurve.OAuthVersion.v2);

var identity.login(function (error) {
	var graph = new Kurve.Graph({ identity: identity });

}
```

<b>Note:</b> In the V2 model, you can pre-define the permission scopes you will need during logon. This avoids causing the app to frequently prompting for permissions during runtime as needed. In other words, if there are permissions you need for sure and want the user to consent them right away, ask them at the logon time:


```javascript

identity.loginAsync([Kurve.Scopes.Mail.Read,
					 Kurve.Scopes.General.OpenId]).then(function(){
                     
                     	}
                     )
```

The scopes are specified in an array. The builtin scopes in the Microsoft Graph are listed under <b>Kurve.Scopes</b>. Any other custom scope needs to be specified as a string. 

Once authentication is done and the graph is created and bound to the identity object, we can then start using the graph objects:

```javascript
//Get the top 5 users
graph.users((function(users, error) {
                getUsersCallback(users, error);
            }), "$top=5");

//Get the top 5 users with Promises
graph.usersAsync("$top=5").then((function(users) {
	getUsersCallback(users, null);

//Get user "me" and then the user's e-mails from Exchange Online:
graph.meAsync().then(function(user)  {
	var result = "User:" + user.data.displayName;
    user.messagesAsync(“$top=5”).then(function(messages) {
		messagesCallback(messages);
	});
});

function messagesCallback(messages) {
	if (messages.nextLink){
    //Check for messages and see if there's
    //another page to be loaded
    	messages.nextLink().then(messagesCallback);
    }
}
```

<b>Note:</b> In the V2 model, Kurve.Graph understands which scopes are required for each operation. Effectively this means that if the user hasn't consented for that permission yet, a prompt for the required consent will activate automatically. The benefit here is that the developer doesn't have to know or code which permissions are required for each of the graph operations.

Just like in the graph example, with identity it is also possible to choose callbacks or Promises syntax:

```javascript
var identity = new Kurve.Identity(
        "your_clientid",
        "your_redirect_url",
        Kurve.OAuthVersion.v2);

identity.login(function(error) {
	if (!error){
	//login worked
    }
});
```

If you prefer the Promises syntax, you can also do it that way:

```javascript
var identity = new Kurve.Identity(
        "your_clientid",
        "your_redirect_url",
        Kurve.OAuthVersion.v2);

identity.loginAsync().then(function() {

});
```
<b>Requesting access tokens:</b>

Differently than what happens with the V1 model where an access token is requested for a given resource, in the V2 model we request access tokens for lists of scopes:

```javascript

  identity.getAccessTokenForScopesAsync([Kurve.Scopes.User.Read, Kurve.Scopes.Mail.Read],false).then(function(token)  {
                document.getElementById("results").innerText = "token: " + token;
            });
```

The false parameter means we want to try getting the access token without a UI prompt. If the user has already consented for that list of scopes, the access token will be returned silently. Otherwise, a UI popup will appear requesting for that consent. If we want to trigger the popup (for example, we know for user a popup is needed and we want to save a roundtrip) then we can set that parameter to true (not recommended). Like in V1, access tokens are cached and consents are remembered next time the user uses the app.

The sample under Examples/OAuthV2 shows how to wire and use it.


## References


### App Model V2

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-compare/ ">https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-compare/ </a>

### Converged Graph API

<a href="http://graph.microsoft.io">http://graph.microsoft.io</a>
 
### Registering an Application on Azure AD

<a href="https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp">https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp</a>

### Incremental Consents

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-scopes/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-scopes/</a>

### Implicit Flow

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-protocols-implicit/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-protocols-implicit/</a>




