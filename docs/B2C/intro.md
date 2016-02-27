# KurveJS and Azure AD B2C - Getting Started

As discussed in the <a href="../../README.md">readme document</a>, Azure AD B2C allows users to sign up/sign in/edit their profiles using external identity proviers. Although from the protocol point of view, B2C and App model V2 are extremely similar, the registration and setup steps are completely different. 

This document will explain how to get started with Kurve JS and AAD B2C.


## Step by step guide

###Step 1: Register your application with AAD B2C and configure the identity providers, policies and attributes.

First, create a B2C AD tenant following these steps: <a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-get-started/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-get-started/</a>

Second, regiter your application following these steps: <a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-app-registration/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-app-registration/</a>

Now you will need to register applications at the identity providers you intend to use. For example, let us assume you will want to authenticate users using their Facebook accounts. For that, you need to register an application at Facebook following these steps: <a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-setup-fb-app/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-setup-fb-app/</a>. Note at the end of the page, you will need to go back to your B2C tenant and create policies for sign up, sign in and edit, attaching the identity provider "Facebook" and configuring it with the settings you just got from Facebook application itself (App ID and App Secret).

Important notes in this step:

1-Different than with the app model v2, there's no authorization flow here. You can request for access tokens to specific resources. What this model will let you do is sign up and sign in users to your app with their identities from other identity providers. That's it.

2-Make sure you have enabled the implicit flow as you register the application under Azure AD B2C. You will need implicit flow for this framework to work.

3-Make sure the redirect URL points to where the login page will be. Below you will see as part of the guidance that you need to copy the sample login.html to your website, so if it points to something like https://localhost:8000/login.html, this will be exactly what the reply URL will also have to be.

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


//Option 2: Automatically linking to the Identity object (note we specify we are working with the app model V2)
var identity = new Kurve.Identity({
                clientId: "your_app_id",
                tokenProcessingUri: "your_redirect_url",
                version: Kurve.OAuthVersion.v2
            });

identity.loginAsync({ policy: "your_policy_name", tenant:"your_tenant_name" }).then(function() {

```

<b>Note:</b> There are two major differences from the standard V2 model here: First, you need a policy name. That policy was defined at your AAD B2C tenant. Also note that a signup policy is different than a signin policy. Therefore, you need to know whether the user is signing up for the first time into your app or just coming back (sign up button vs login button). That is why you will want to define the policy at the login method, not the constructor. Second, you need to specify your tenant name from Azure AD B2C. For example, in the published sample I have "matb2c.onmicrosoft.com". If you don't specify a tenant, we would have to use "common" (which is what we normally use for other operations) and that won't be allowed for B2C scenarios. In other words, if you try specifying a policy (which means you want to use B2C) but not a tenant name, we will return an error.

Once you are authenticated, you can ask for details about the token:


```javascript
JSON.stringify(identity.getIdToken())

```

That will give you everything we know from that user that B2C has given us. You can define attributes you are looking for in your B2C tenant configuration, such as display name, job title, email, etc.

As mentioned before, there is no Kurve.Graph in the B2C model as these users won't have Graph API access.



## References


### App Model V2

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-compare/ ">https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-compare/ </a>

### Tokens in AAD B2C

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-reference-tokens/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-reference-tokens/</a>

### OpenID Connect reference for B2C

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-reference-oidc/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-reference-oidc/</a>

### Policies in B2C

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-reference-policies/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-reference-policies/</a>

### Getting Started with B2C

<a href="https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-get-started/">https://azure.microsoft.com/en-us/documentation/articles/active-directory-b2c-get-started/</a>




