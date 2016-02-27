# KurveJS and App model V1 - Getting Started

As discussed in the <a href="../../README.md">readme document</a>, the app model v1 is the currently supported model for production scenaarios.

This document will explain how to get started with Kurve JS and this model.


## Step by step guide

###Step 1: Register your application in Azure Active Directory.

Your application needs an ID. In order to get one, you need to go to <a href="https://manage.windowsazure.com">https://manage.windowsazure.com</a>. Under active directory, register a web application and set it up for implicit grant flow. If you are unsure how to do that, follow this topic (remember to also follow the implicit grant flow section in there): <a href="https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp">https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp</a>

Important notes in this step:

1-Make sure you also set the permissions your application will need. For example, in order to use the Kurve.Graph class, you will need to add the "Microsoft Graph" in the "permissions to other applications" section and then add the corresponding permissions the app will actually need

2-Make sure you have enabled the implicit grant flow as mentioned above, otherwise Kurve.JS won't work

3-Make sure the reply URL points to where the login page will be. Below you will see as part of the guidance that you need to copy the sample login.html to your website, so if it points to something like https://localhost:8000/login.html, this will be exactly what the reply URL will also have to be.

###Step 2: Reference the JavaScript file from your HTML page:

```html
    <script src="Kurve.min-0.4.0.js"></script>
```

Alternatively you can just reference the script from our CDN:

```html
    <script src="https://kurvejs.blob.core.windows.net/dist/kurve.min-0.4.0.js"></script>
```

###Step 3: Code:


The identity and graph classes can work independently or together. For example, you may have authenticated and retrieved an access token from somewhere else and only give it to the graph library, or you can use the identity library for authentication.

```javascript

//Option 1: We already have an access token obtained from another library:
var graph = new Kurve.Graph({ defaultAccessToken: token });


//Option 2: Automatically linking to the Identity object (note we specify we are working with the app model V1)
var identity = new Kurve.Identity({
                clientId: "your_app_id",
                tokenProcessingUri: "your_redirect_url",
                version: Kurve.OAuthVersion.v1
            });

var identity.login(function (error) {
	var graph = new Kurve.Graph({ identity: identity });

}
```

And then we can start using the graph objects:

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

Just like in the graph example, with identity it is also possible to choose callbacks or Promises syntax:

```javascript
var identity = new Kurve.Identity({
                clientId: "your_app_id",
                tokenProcessingUri: "your_redirect_url",
                version: Kurve.OAuthVersion.v1
            });

identity.login(function(error) {
	if (!error){
	//login worked
    }
	identity.getAccessToken("https://graph.microsoft.com",(function (token,error) {
    	if (!error){
		//token received
        }
	});
});
```

If you prefer the Promises syntax, you can also do it that way:

```javascript
var identity = new Kurve.Identity({
                clientId: "your_app_id",
                tokenProcessingUri: "your_redirect_url",
                version: Kurve.OAuthVersion.v1
            });

identity.loginAsync().then(function() {
	identity.getAccessTokenAsync("https://graph.microsoft.com").then(function (token) {
		//token received
	});
});
```

The sample under Examples/OAuthV1 shows how to wire and use it.


## References

### Converged Graph API

<a href="http://graph.microsoft.io">http://graph.microsoft.io</a>
 
### Registering an Application on Azure AD

<a href="https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp">https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp</a>

