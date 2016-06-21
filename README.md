# KurveJS

Kurve<nolink>.JS is an experimental, unofficial, open source JavaScript / TypeScript library that aims to provide:

* Easy authentication and authorization using Azure Active Directory.
* Easy access to the Microsoft Graph REST API, utilizing Intellisense in your TypeScript-aware editor

It looks like this:

```typescript
const id = new kurve.Identity("your-client-id", "http://your/path/to/login.html");

id.loginAsync().then(_ => {
    const graph = new kurve.Graph(id);

    graph.me.messages.GetMessages().then(messages => {
        console.log("my emails:");   
        messages.value.forEach(message => 
            console.log(message.subject)
        )
    });
    
    graph.me.manager.GetUser(manager => {
        console.log("my manager", manager.displayName);
        console.log("their directs:");
        manager._context.directReports.GetUsers(directs =>
            directs.value.forEach(direct =>
                console.log(direct.displayName)
            )
        )
    });
});
```

Kurve works well with most JavaScript and TypeScript frameworks including Angular 1, Angular 2, Ember, and React.

Kuve enables developers building web applications - including single application pages - to support a range of authentication and authorization scenarios including:

* AAD app model v1 with a postMessage flow
* AAD app model v1 with redirections
* AAD app model v2 with a postMessage flow
* AAD app model v2 with Active Directoy B2C.  This makes it easy to add third party identity providers such as Facebook and signup/signin/profile edit experiences using AD B2C policies

## Setup

kurve.js is a UMD file, allowing maximum flexibility.

### node

_Note: node support is currently highly experimental._ 

Install kurve from npm:

```
npm install kurvejs
```

Include kurve.js into your project:

```typescript
import kurve = require ("kurvejs"); // typescript
var kurve = require("kurvejs");     // javascript
```
    
### browser via npm

Install kurve from npm:

```
npm install kurvejs
```
    
Copy login.html to your source tree.

Include kurve.js into your project:

```typescript
import kurve = require ("kurvejs"); // typescript
var kurve = require("kurvejs");     // javascript
```

Bundle kurve into your app using webpack, browserify, etc.

### browser via &lt;script/&gt;

Include kurve.js into your html:

```html
<script src="kurve.js"/>
```

Copy login.html to your source tree.

If using TypeScript, add the following:

```typescript
/// <reference path="kurve-global.d.ts"/>  
const kurve = window["Kurve"] as typeof Kurve; 
```

Note that _Kurve_ (capitalized) is the namespace that declares the types, where _kurve_ (lowercase) is the module. Use _Kurve_ when declaring types, and _kurve_ for everything else: 

```typescript
const error:Kurve.Error = new kurve.Error()
```

## Building Kurve

```
npm install
npm run build
```

## Using Kurve Identity

The first thing you have to decide before using Kurve is which app model version you will use:

### App Model V1:

App model V1 is the current, production supported model in Azure AD. This model implies that:

1. You will register an application in Azure Active Directory, using the Azure Management Console https://manage.windowsazure.com
2. During that registration, you will pre-set the permissions your application will need. This means that you won't make those decisions about what types of permissions an app needs in runtime, but instead predefine those when you register your application in Azure AD
3. Every access token will be requested againast a resource, not a scope. For example, a resource could be Microsoft Exchange Online and another would be your custom web service

### App Model V2:

App model V2 is a new, more modern way which is still in Preview and has limitations at this point. It enables more flexible scenarios, including AAD B2C which leverages external identity providers such as Facebook and user signup/signin/profile edit support. This model implies that:

1. You will register an application in the new application registration portal under https://apps.dev.microsoft.com, which does not require using Azure Management Portal.
2. You do not specify application permissions during the registration. The application code itself will request for specific permissions in runtime (either during login or any time during the application execution). Kuve JS will help with that process.

### App Model V2 with AAD B2C:

As mentioned above, B2C enables users to sign up to your AAD tenant using external identity providers such as Facebook, Google, LinkedIn, Amazon and Microsoft Acount. You define policies and the attributes you want to collect during the sign up, for example a user might have to enter their name, e-mail and phone number so that gets recorded into your tenant and accessible to your application.

1. Create a B2C AD tenant following these steps: https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-get-started
2. Register your application following these steps: https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-app-registration
3. Now you will need to register applications at the identity providers you intend to use. For example, let us assume you will want to authenticate users using their Facebook accounts. For that, you need to register an application at Facebook following these steps: https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-setup-fb-app/. Note at the end of the page, you will need to go back to your B2C tenant and create policies for sign up, sign in and edit, attaching the identity provider "Facebook" and configuring it with the settings you just got from Facebook application itself (App ID and App Secret).

Important notes in this step:

1. Different than with the app model v2, there's no authorization flow here. You can request for access tokens to specific resources. What this model will let you do is sign up and sign in users to your app with their identities from other identity providers. That's it.

2. Make sure you have enabled the implicit flow as you register the application under Azure AD B2C. You will need implicit flow for this framework to work.

3. Make sure the redirect URL points to where the login page will be. Below you will see as part of the guidance that you need to copy the sample login.html to your website, so if it points to something like https://localhost:8000/login.html, this will be exactly what the reply URL will also have to be.

## Using Kurve Graph

You can use Kurve Graph without Kurve Identity by passing an access token directly:

```typescript
const graph = new kurve.Graph("access_token");
```

Or else link it to a Kurve.Identity object:

```typescript
const graph = new kurve.Graph(idKurve);
```
    
For information on accessing the graph, see the [QueryBuilder documentation](./graph.md).

## Samples

To run samples, start an http server using port 8000 in the root directory of this repository and aim your browser at http://localhost:port/samples.html

## FAQ

### Is this a supported library from Microsoft?

No it is not. This is an experimental unofficial open source project. If you are looking for a supported APIs we encourage you to call Microsoft's Graph REST APIs directly.

### Can I use/change this library?

You are free to take the code, change and use it any way you want it. But please be advised this code isn't supported.

### What if I find issues or want to contribute/help?

You are free to send us your feedback at this Github repo, send pull requests, etc. But please don't expect this to work as an official support channel

# Release Notes

## 0.6.0
 * Revised Identity & Graph constructors with better defaults
 * Separate type definitions file for global (&lt;script>) module import  

## 0.5.0
 * Refactored source tree
 * Initial node support (highly experimental)
 * Build using "npm run build" instead of Gulp
 * New Graph access via QueryBuilder
 * Updated documentation & samples

## 0.4.2:
 * Cached tokens can now be persisted to a local store
 * Expanded Graph support including event, mailFolders, messageAttachments
 * Simplified code
 * Bug fixes

## 0.4.1:
 * Added support for /calendarView calendar events
 * Simplified code with inheritance and generics
 * Bug fixes

## 0.4.0:
 * Added support for AAD B2C
 * Breaking change: identity constructor now uses an object to group parameters
 * Breaking change: login method now uses an object to group parameters
 * By default now in app model v2, both openid connect and profile scopes are requested
 * Bug fixes, more examples

## 0.3.7:
 * Added support for calendar events

## 0.3.6:
 * Bug fixes
 * Introduction of the app model V2 support, including dynamic scopes and integration with the graph calls
 * Introduction of no window support, for older browsers that don't support popup windows during login

## 0.3.5:
 * New typed promises syntax supporting return and exception types
 
## 0.3.1:
 * Hotfix for the token expiration loop issue

## 0.3.0:
 * Better type support for all returned types in callbacks
 * Standardized return entities and collections into .data for all related properties returned from the graph so we differentiate the model data from action methods more easily

## 0.2.0:
 * Minification and unification of the library into a single file
 * No need for tenant ID anymore
 * Improved error handling, all error returns are now coming into the Kurve.Error class format
 * Additional graph methods, like reading groups, user profile's photos, user manager and direct reports

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
