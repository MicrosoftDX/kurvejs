# KurveJS

Kurve<nolink>.JS is an unofficial, open source JavaScript / TypeScript library that aims to simplify two tasks:

1. Easy authentication and authorization against Azure Active Directory - without navigating away from you web app for login or token requests.

2. Easy access to the Microsoft Graph REST API.

## How does it work?

Reference the JavaScript file from your HTML page:

```html
    <script src="Kurve.min.js"></script>
```

Alternatively you can just reference the script from our CDN:

```html
    <script src="https://kurvejs.blob.core.windows.net/dist/Kurve.min-0.2.0.js"></script>
```

	
The identity and graph classes can work independently or together. For example, you may have authenticated and retrieved an access token from somewhere else and only give it to the graph library, or you can use the identity library for authentication.

```javascript

//Option 1: We already have an access token obtained from another library: 
var graph = new Kurve.Graph({ defaultAccessToken: token });


//Option 2: Automatically linking to the Identity object
var identity = new Kurve.Identity(
        "your_client_id", 
        "your_redirect_url");

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
	var result = "User:" + user.displayName;
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
var identity = new Kurve.Identity(
        "your_clientid", 
        "your_redirect_url");

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
var identity = new Kurve.Identity(
        "your_clientid", 
        "your_redirect_url");

identity.loginAsync().then(function() {
	identity.getAccessTokenAsync("https://graph.microsoft.com").then(function (token) {
		//token received
	});
});
```

The sample index.html and app.<nolink>js files show how to wire and use it.


## FAQ

### Is this a supported library from Microsoft?
 
No it is not. This is an open source project built unofficially. If you are looking for a supported APIs we encourage you to directly call Microsoft's Graph REST APIs. 
 
### Can I use/change this library?

You are free to take the code, change and use it any way you want it. But please be advised this code isn't supported.

### What if I find issues or want to contribute/help?

You are free to send us your feedback at this Github repo, send pull requests, etc. But please don't expect this to work as an official support channel

### Which files do I need to run this? 

At minimum you need the KurveGraph.<nolink>js and Promises.<nolink>js, and optionally KurveIdentity.<nolink>js + login.html. You may use the TypeScript libraries and reuse some of the sample app code (index.html and app.<nolink>js) for reference.

# Release Notes

## 0.2.0:
 * Minification and unification of the library into a single file
 * No need for tenant ID anymore
 * Improved error handling, all error returns are now coming into the Kurve.Error class format
 * Additional graph methods, like reading groups, user profile's photos, user manager and direct reports
