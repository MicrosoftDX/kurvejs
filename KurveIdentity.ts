// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {

    export enum OAuthVersion {
        v1=1,
        v2=2
    }
    export class Error {
        public status: number;
        public statusText: string;
        public text: string;
        public other: any;
    }
    class Token {
        id: string;
        scopes: string[];
        resource: string;
        token: string;
        expiry: Date;
    }
    interface TokenDictionary {
        [index: string]: Token;
    }

    export class IdToken {
        public Token: string;
        public IssuerIdentifier: string;
        public SubjectIdentifier: string;
        public Audience: string;
        public Expiry: Date;
        public UPN: string;
        public TenantId: string;
        public FamilyName: string;
        public GivenName: string;
        public Name: string;
        public PreferredUsername: string;

    }
    
     export class Identity {
        public authContext: any = null;
        public config: any = null;
        public isCallback: boolean = false;
        public clientId: string;
        private req: XMLHttpRequest;
        private state: string;
        private version: OAuthVersion;
        private nonce: string;
        private idToken: IdToken;
        private loginCallback: (error: Error) => void;
        private accessTokenCallback: (token:string, error: Error) => void;
        private getTokenCallback: (token: string, error: Error) => void;
        private redirectUri: string;
        private tokenCache: TokenDictionary;
        private logonUser: any;
        private refreshTimer: any;

        constructor(clientId = "", redirectUri = "", version?: OAuthVersion) {
  
            this.clientId = clientId;
            this.redirectUri = redirectUri;
            this.req = new XMLHttpRequest();
            this.tokenCache = {};
            if (version)
                this.version = version;
            else
                this.version = OAuthVersion.v1;

            //Callback handler from other windows
            window.addEventListener("message", ((event: MessageEvent) => {
                if (event.data.type === "id_token") {
                    if (event.data.error) {
                        var e: Error = new Error();
                        e.text = event.data.error;
                        this.loginCallback(e);
                        
                    } else {
                        //check for state
                        if (this.state !== event.data.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            this.loginCallback(error);
                        } else {
                            this.decodeIdToken(event.data.token);
                            this.loginCallback(null);
                        }
                    }
                } else if (event.data.type === "access_token") {
                    if (event.data.error) {
                        var e: Error = new Error();
                        e.text = event.data.error;
                        this.getTokenCallback(null, e);

                    } else {
                        var token:string = event.data.token;
                        var iframe = document.getElementById("tokenIFrame");
                        iframe.parentNode.removeChild(iframe);

                        if (event.data.state !== this.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            this.getTokenCallback(null, error);
                        }
                        else {
                            this.getTokenCallback(token, null);
                        }
                    }
                }
            }));
        }

        private decodeIdToken(idToken: string): void {
           
            var decodedToken = this.base64Decode(idToken.substring(idToken.indexOf('.') + 1, idToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            this.idToken = new IdToken();
            this.idToken.Token = idToken;
            this.idToken.Expiry = expiryDate;
            this.idToken.UPN = decodedTokenJSON.upn;
            this.idToken.TenantId = decodedTokenJSON.tid;
            this.idToken.FamilyName = decodedTokenJSON.family_name;
            this.idToken.GivenName = decodedTokenJSON.given_name;
            this.idToken.Name = decodedTokenJSON.name;
            this.idToken.PreferredUsername = decodedTokenJSON.preferred_username;

            
            var expiration: Number = expiryDate.getTime() - new Date().getTime() - 300000;

            this.refreshTimer = setTimeout((() => {
                this.renewIdToken();
            }), expiration); 
        }

        private decodeAccessToken(accessToken: string, resource?:string, scopes?:string[]): void {
            var decodedToken = this.base64Decode(accessToken.substring(accessToken.indexOf('.') + 1, accessToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            var key: string;
            if (resource)
                key = resource;
            else
                key = scopes.join(" ");
            var token = new Token();
            token.expiry = expiryDate;
            token.resource = resource;
            token.scopes = scopes;
            token.token = accessToken;
            token.id = key;

            this.tokenCache[key] = token;
        }

        public getIdToken(): any {
            return this.idToken;
        }
        public isLoggedIn(): boolean {
            if (!this.idToken) return false;
            return (this.idToken.Expiry > new Date());
        }

        private renewIdToken(): void {
            clearTimeout(this.refreshTimer);
            this.login((() => {

            }));
        }

        public getCurrentOauthVersion(): OAuthVersion {
            return this.version;
        }

        public getAccessTokenAsync(resource: string): Promise<string,Error> {

            var d = new Deferred<string,Error>();
            this.getAccessToken(resource, ((token, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        }   

        public getAccessToken(resource: string, callback: (token: string, error: Error) => void): void {
            if (this.version !== OAuthVersion.v1) {
                var e = new Error();
                e.statusText = "Currently this identity class is using v2 OAuth mode. You need to use getAccessTokenForScopes() method";
                callback(null, e);
                return;
            }
            //Check for cache and see if we have a valid token
            var cachedToken = null;
            var keys = Object.keys(this.tokenCache);
            keys.forEach((key) => {
                var token = this.tokenCache[key];
                
                //remove expired tokens
                if (token.expiry <= (new Date(new Date().getTime() + 60000))) {
                    delete this.tokenCache[key];
                } else {
                    //Tries to capture a token that matches the resource
                    var containScopes = true;
                    if (token.resource == resource) {
                        cachedToken = token;
                    }
                }
            });

            if (cachedToken) {
                //We have it cached, has it expired? (5 minutes error margin)
                if (cachedToken.expiry > (new Date(new Date().getTime() + 60000))) {
                    callback(<string>cachedToken.token, null);
                    return;
                }
            }
            //If we got this far, we need to go get this token

            //Need to create the iFrame to invoke the acquire token
            this.getTokenCallback = ((token: string, error: Error) => {
                if (error) {
                    callback(null, error);
                }
                else {
                    this.decodeAccessToken(token, resource);
                    callback(token, null);

                }
            });

            this.nonce = "token" + this.generateNonce();
            this.state = "token" + this.generateNonce();

            var iframe = document.createElement('iframe');
            iframe.style.display = "none";
            iframe.id = "tokenIFrame";
            
            iframe.src = this.redirectUri + "?clientId=" + encodeURIComponent(this.clientId) +
            "&resource=" + encodeURIComponent(resource) +
            "&redirectUri=" + encodeURIComponent(this.redirectUri) +
                "&state=" + encodeURIComponent(this.state) +
                "&version=" + encodeURIComponent(this.version.toString()) +
                "&nonce=" + encodeURIComponent(this.nonce) +
                "&op=token";
            document.body.appendChild(iframe);
        }

        public getAccessTokenForScopesAsync(scopes: string[], promptForConsent = false): Promise<string, Error> {
            var d = new Deferred<string, Error>();
            this.getAccessTokenForScopes(scopes, promptForConsent, ((token, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        }   

        public getAccessTokenForScopes(scopes: string[], promptForConsent=false, callback: (token: string, error: Error) => void): void {
            if (this.version !== OAuthVersion.v2) {
                var e = new Error();
                e.statusText = "Dynamic scopes require v2 mode. Currently this identity class is using v1";
                callback(null, e);
                return;
            }
            
            //Check for cache and see if we have a valid token
            var cachedToken = null;
            var keys = Object.keys(this.tokenCache);
            keys.forEach((key) => {
                var token = this.tokenCache[key];
                
                //remove expired tokens
                if (token.expiry <= (new Date(new Date().getTime() + 60000))) {
                    delete this.tokenCache[key];
                } else {
                    //Tries to capture a token that contains all scopes and is still valid
                    var containScopes = true;
                    if (token.scopes) {
                        scopes.forEach((scope: string) => {
                            if (token.scopes.indexOf(scope) < 0)
                                containScopes = false;
                        });
                    }

                    if (containScopes) {
                        cachedToken = token;
                    }
                }
            });
            if (cachedToken) {
                //We have it cached, has it expired? (5 minutes error margin)
                if (cachedToken.expiry > (new Date(new Date().getTime() + 60000))) {
                    callback(<string>cachedToken.token, null);
                    return;
                }
            }

            //If we got this far, we don't have a valid cached token, so will need to get one.

            //Need to create the iFrame to invoke the acquire token

            this.getTokenCallback = ((token: string, error: Error) => {
                if (error) {
                    if (promptForConsent || !error.text) {
                        callback(null, error);
                    } else if (error.text.indexOf("AADSTS65001")>=0) {
                        //We will need to try getting the consent
                        this.getAccessTokenForScopes(scopes, true, this.getTokenCallback);
                    } else {
                        callback(null, error);
                    }
                }
                else {
                    this.decodeAccessToken(token, null, scopes);
                    callback(token, null);
                }
            });

            this.nonce = "token" + this.generateNonce();
            this.state = "token" + this.generateNonce();

            if (!promptForConsent) {
                var iframe = document.createElement('iframe');
                iframe.style.display = "none";
                iframe.id = "tokenIFrame";
                iframe.src = this.redirectUri + "?clientId=" + encodeURIComponent(this.clientId) +
                    "&scopes=" + encodeURIComponent(scopes.join(" ")) +
                    "&redirectUri=" + encodeURIComponent(this.redirectUri) +
                    "&version=" + encodeURIComponent(this.version.toString()) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce) +
                    "&login_hint=" + encodeURIComponent(this.idToken.PreferredUsername) +
                    "&domain_hint=" + encodeURIComponent(this.idToken.TenantId === "9188040d-6c67-4c5b-b112-36a304b66dad" ? "consumers" : "organizations") +
                    "&op=token";
                document.body.appendChild(iframe);
            } else {
                window.open(this.redirectUri + "?clientId=" + encodeURIComponent(this.clientId) +
                    "&scopes=" + encodeURIComponent(scopes.join(" ")) +
                    "&redirectUri=" + encodeURIComponent(this.redirectUri) +
                    "&version=" + encodeURIComponent(this.version.toString()) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce) +
                    "&op=token" 
                    , "_blank");
            }


        }

        public loginAsync(scopes?:string[]): Promise<void, Error> {
            var d = new Deferred<void,Error>();
            this.login(((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }), scopes);
            return d.promise;
        }

        public login(callback: (error: Error) => void, scopes?:string[]): void {
            this.loginCallback = callback;
          
            if (scopes && this.version === OAuthVersion.v1) {
                var e = new Error();
                e.text = "Scopes can only be used with OAuth v2.";
                callback(e);
                return;
            }
            this.state = "login" + this.generateNonce();
            this.nonce = "login" + this.generateNonce();

            var loginURL = this.redirectUri + "?clientId=" + encodeURIComponent(this.clientId) +
                "&redirectUri=" + encodeURIComponent(this.redirectUri) +
                "&state=" + encodeURIComponent(this.state) +
                "&nonce=" + encodeURIComponent(this.nonce) +
                "&version=" + encodeURIComponent(this.version.toString()) +
                "&op=login";

            if (scopes) {
                loginURL += "&scopes=" + encodeURIComponent(scopes.join(" "));
            }

            
            window.open(loginURL, "_blank");
        }

        public logOut(): void {
            var url = "https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=" + encodeURI(window.location.href);
            window.location.href = url;
        }

        private base64Decode(encodedString: string): string {
            var e = {}, i, b = 0, c, x, l = 0, a, r = '', w = String.fromCharCode, L = encodedString.length;
            var A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            for (i = 0; i < 64; i++) { e[A.charAt(i)] = i; }
            for (x = 0; x < L; x++) {
                c = e[encodedString.charAt(x)]; b = (b << 6) + c; l += 6;
                while (l >= 8) { ((a = (b >>> (l -= 8)) & 0xff) || (x < (L - 2))) && (r += w(a)); }
            }
            return r;
        }

        private generateNonce(): string {
            var text = "";
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            for (var i = 0; i < 32; i++) {
                text += chars.charAt(Math.floor(Math.random() * chars.length));
            }
            return text;
        }
    }
}


//*********************************************************   
//   
//Kurve js, https://github.com/microsoftdx/kurvejs
//  
//Copyright (c) Microsoft Corporation  
//All rights reserved.   
//  
// MIT License:  
// Permission is hereby granted, free of charge, to any person obtaining  
// a copy of this software and associated documentation files (the  
// ""Software""), to deal in the Software without restriction, including  
// without limitation the rights to use, copy, modify, merge, publish,  
// distribute, sublicense, and/or sell copies of the Software, and to  
// permit persons to whom the Software is furnished to do so, subject to  
// the following conditions:  




// The above copyright notice and this permission notice shall be  
// included in all copies or substantial portions of the Software.  




// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,  
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND  
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE  
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION  
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
//   
//*********************************************************   
