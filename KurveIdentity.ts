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

    class CachedToken {
        id: string;
        scopes: string[];
        resource: string;
        token: string;
        expiry: Date;

        constructor(token: CachedToken);
        constructor(token?: { id?: string, scopes?: string[], resource?: string, token?: string, expiry?: string });
        constructor(token: any) {
            token = token || {};
            this.id = token.id,
            this.scopes = token.scopes;
            this.resource = token.resource;
            this.token = token.token;
            this.expiry = new Date(token.expiry);
        }

        public get isExpired() {
            return this.expiry <= new Date(new Date().getTime() + 60000);
        }

        public hasScopes(requiredScopes) {
            var actualScopes = <Array<string>>this.scopes || [];
            return requiredScopes.every(requiredScope => {
                return actualScopes.some(actualScope => requiredScope == actualScope);
            });
        }
    }

    interface CachedTokenDictionary {
        [index: string]: CachedToken;
    }

    export interface TokenStorage {
        add(key: string, token: any);
        remove(key: string);
        getAll(): any[];
        clear();
    }

    class TokenCache {
        private cachedTokens: CachedTokenDictionary;

        constructor(private tokenStorage: TokenStorage) {
            this.cachedTokens = {};
            if (tokenStorage) {
                tokenStorage.getAll().forEach((token) => {
                    token = new CachedToken(token);
                    if (token.isExpired) {
                        this.tokenStorage.remove(token);
                    } else {
                        this.cachedTokens[token.id] = token;
                    }
                });
            }
        }

        public add(token: CachedToken) {
            this.cachedTokens[token.id] = token;
            this.tokenStorage && this.tokenStorage.add(token.id, token);
        }

        public getForResource(resource: string): CachedToken {
            var cachedToken = this.cachedTokens[resource];
            if (cachedToken && cachedToken.isExpired) {
                this.remove(resource);
                return null;
            }
            return cachedToken;
        }

        public getForScopes(scopes: string[]): CachedToken {
            for (var key in this.cachedTokens) {
                var cachedToken = this.cachedTokens[key];

                if (cachedToken.hasScopes(scopes)) {
                    if (cachedToken.isExpired) {
                        this.remove(key);
                    } else {
                        return cachedToken;
                    }
                }
            }

            return null;
        }

        public clear() {
            this.cachedTokens = {};
            this.tokenStorage && this.tokenStorage.clear();
        }

        private remove(key) {
            this.tokenStorage && this.tokenStorage.remove(key);
            delete this.cachedTokens[key];
        }
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
        public FullToken: any;

    }

    export interface IdentitySettings {
        clientId: string;
        tokenProcessingUri: string;
        version: OAuthVersion;
        tokenStorage?: TokenStorage;
    }

    export class Identity {
//      public authContext: any = null;
//      public config: any = null;
//      public isCallback: boolean = false;
        public clientId: string;
//      private req: XMLHttpRequest;
        private state: string;
        private version: OAuthVersion;
        private nonce: string;
        private idToken: IdToken;
        private loginCallback: (error: Error) => void;
//      private accessTokenCallback: (token: string, error: Error) => void;
        private getTokenCallback: (token: string, error: Error) => void;
        private tokenProcessorUrl: string;
        private tokenCache: TokenCache;
//      private logonUser: any;
        private refreshTimer: any;
        private policy: string = "";
//      private tenant: string = "";

        constructor(identitySettings: IdentitySettings) {
            this.clientId = identitySettings.clientId;
            this.tokenProcessorUrl = identitySettings.tokenProcessingUri;
//          this.req = new XMLHttpRequest();
            if (identitySettings.version)
                this.version = identitySettings.version;
            else
                this.version = OAuthVersion.v1;

            this.tokenCache = new TokenCache(identitySettings.tokenStorage);

            //Callback handler from other windows
            window.addEventListener("message", event => {
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
            });
        }

        public checkForIdentityRedirect(): boolean {
            function token(s: string) {
                var start = window.location.href.indexOf(s);
                if (start < 0) return null;
                var end = window.location.href.indexOf("&",start + s.length);
                return  window.location.href.substring(start,((end > 0) ? end : window.location.href.length));
            }

            function parseQueryString(str: string) {
                var queryString = str || window.location.search || '';
                var keyValPairs: any[] = [];
                var params: any = {};
                queryString = queryString.replace(/.*?\?/, "");

                if (queryString.length) {
                    keyValPairs = queryString.split('&');
                    for (var pairNum in keyValPairs) {
                        var key = keyValPairs[pairNum].split('=')[0];
                        if (!key.length) continue;
                        if (typeof params[key] === 'undefined')
                            params[key] = [];
                        params[key].push(keyValPairs[pairNum].split('=')[1]);
                    }
                }
                return params;
            }

            var params = parseQueryString(window.location.href);
            var idToken = token("#id_token=");
            var accessToken = token("#access_token");
            if (idToken) {
                if (true || this.state === params["state"][0]) { //BUG? When you are in a pure redirect system you don't remember your state or nonce so don't check.
                    this.decodeIdToken(idToken);
                    this.loginCallback && this.loginCallback(null);
                } else {
                    var error = new Error();
                    error.statusText = "Invalid state";
                    this.loginCallback && this.loginCallback(error);
                }
                return true;
            }
            else if (accessToken) {
                throw "Should not get here.  This should be handled via the iframe approach."
/*
                if (this.state === params["state"][0]) {
                    this.getTokenCallback && this.getTokenCallback(accessToken, null);
                } else {
                    var error = new Error();
                    error.statusText = "Invalid state";
                    this.getTokenCallback && this.getTokenCallback(null, error);
                }
*/
            }
            return false;
        }

        private decodeIdToken(idToken: string): void {

            var decodedToken = this.base64Decode(idToken.substring(idToken.indexOf('.') + 1, idToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            this.idToken = new IdToken();
            this.idToken.FullToken = decodedTokenJSON;
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
            var key = resource || scopes.join(" ");
            var token = new CachedToken();
            token.expiry = expiryDate;
            token.resource = resource;
            token.scopes = scopes;
            token.token = accessToken;
            token.id = key;

            this.tokenCache.add(token);
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
            this.login(() => { });
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

        public getAccessToken(resource: string, callback: PromiseCallback<string>): void {
            if (this.version !== OAuthVersion.v1) {
                var e = new Error();
                e.statusText = "Currently this identity class is using v2 OAuth mode. You need to use getAccessTokenForScopes() method";
                callback(null, e);
                return;
            }

            var token = this.tokenCache.getForResource(resource);
            if (token) {
                return callback(token.token, null);
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

            iframe.src = this.tokenProcessorUrl + "?clientId=" + encodeURIComponent(this.clientId) +
            "&resource=" + encodeURIComponent(resource) +
                "&redirectUri=" + encodeURIComponent(this.tokenProcessorUrl) +
                "&state=" + encodeURIComponent(this.state) +
                "&version=" + encodeURIComponent(this.version.toString()) +
                "&nonce=" + encodeURIComponent(this.nonce) +
                "&op=token";

            document.body.appendChild(iframe);
        }


        public getAccessTokenForScopesAsync(scopes: string[], promptForConsent = false): Promise<string, Error> {

            var d = new Deferred<string, Error>();
            this.getAccessTokenForScopes(scopes, promptForConsent, (token, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            });
            return d.promise;
        }

        public getAccessTokenForScopes(scopes: string[], promptForConsent=false, callback: (token: string, error: Error) => void): void {
            if (this.version !== OAuthVersion.v2) {
                var e = new Error();
                e.statusText = "Dynamic scopes require v2 mode. Currently this identity class is using v1";
                callback(null, e);
                return;
            }

            var token = this.tokenCache.getForScopes(scopes);
            if (token) {
                return callback(token.token, null);
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
                iframe.src = this.tokenProcessorUrl + "?clientId=" + encodeURIComponent(this.clientId) +
                    "&scopes=" + encodeURIComponent(scopes.join(" ")) +
                    "&redirectUri=" + encodeURIComponent(this.tokenProcessorUrl) +
                    "&version=" + encodeURIComponent(this.version.toString()) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce) +
                    "&login_hint=" + encodeURIComponent(this.idToken.PreferredUsername) +
                    "&domain_hint=" + encodeURIComponent(this.idToken.TenantId === "9188040d-6c67-4c5b-b112-36a304b66dad" ? "consumers" : "organizations") +
                    "&op=token";
                document.body.appendChild(iframe);
            } else {
                window.open(this.tokenProcessorUrl + "?clientId=" + encodeURIComponent(this.clientId) +
                    "&scopes=" + encodeURIComponent(scopes.join(" ")) +
                    "&redirectUri=" + encodeURIComponent(this.tokenProcessorUrl) +
                    "&version=" + encodeURIComponent(this.version.toString()) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce) +
                    "&op=token"
                    , "_blank");
            }
        }

        public loginAsync(loginSettings?: { scopes?: string[], policy?:string, tenant?:string}): Promise<void, Error> {
            var d = new Deferred<void, Error>();
            this.login((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }, loginSettings);
            return d.promise;
        }

        public login(callback: (error: Error) => void, loginSettings?: { scopes?: string[], policy?: string, tenant?: string}): void {
            this.loginCallback = callback;
            if (!loginSettings) loginSettings = {};
            if (loginSettings.policy) this.policy = loginSettings.policy;

            if (loginSettings.scopes && this.version === OAuthVersion.v1) {
                var e = new Error();
                e.text = "Scopes can only be used with OAuth v2.";
                callback(e);
                return;
            }

            if (loginSettings.policy && !loginSettings.tenant) {
                var e = new Error();
                e.text = "In order to use policy (AAD B2C) a tenant must be specified as well.";
                callback(e);
                return;
            }
            this.state = "login" + this.generateNonce();
            this.nonce = "login" + this.generateNonce();
            var loginURL = this.tokenProcessorUrl + "?clientId=" + encodeURIComponent(this.clientId) +
                "&redirectUri=" + encodeURIComponent(this.tokenProcessorUrl) +
                "&state=" + encodeURIComponent(this.state) +
                "&nonce=" + encodeURIComponent(this.nonce) +
                "&version=" + encodeURIComponent(this.version.toString()) +
                "&op=login" +
                "&p=" + encodeURIComponent(this.policy);
            if (loginSettings.tenant) {
                loginURL += "&tenant=" + encodeURIComponent(loginSettings.tenant);
            }
            if (this.version === OAuthVersion.v2) {
                    if (!loginSettings.scopes) loginSettings.scopes = [];
                    if (loginSettings.scopes.indexOf("profile") < 0)
                        loginSettings.scopes.push("profile");
                    if (loginSettings.scopes.indexOf("openid") < 0)
                        loginSettings.scopes.push("openid");

                    loginURL += "&scopes=" + encodeURIComponent(loginSettings.scopes.join(" "));
            }
            window.open(loginURL, "_blank");
        }


        public loginNoWindowAsync(toUrl? : string): Promise<void, Error> {
            var d = new Deferred<void, Error>();
            this.loginNoWindow((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }, toUrl);
            return d.promise;
        }

        public loginNoWindow(callback: (error: Error) => void, toUrl? : string): void {
            this.loginCallback = callback;
            this.state = "clientId=" + this.clientId + "&" + "tokenProcessorUrl=" + this.tokenProcessorUrl
            this.nonce = this.generateNonce();

            var redirected = this.checkForIdentityRedirect();
            if (!redirected) {
                var redirectUri = (toUrl) ? toUrl : window.location.href.split("#")[0];  // default the no login window scenario to return to the current page
                var url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token" +
                    "&client_id=" + encodeURIComponent(this.clientId) +
                    "&redirect_uri=" + encodeURIComponent(redirectUri) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce);
                window.location.href = url;
            }
        }

        public logOut(): void {
            this.tokenCache.clear();
            var url = "https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=" + encodeURI(window.location.href);
            window.location.href = url;
        }

        private base64Decode(encodedString: string): string {
            var e: any = {}, i: number, b = 0, c: number, x: number, l = 0, a: any, r = '', w = String.fromCharCode, L = encodedString.length;
            var A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            for (i = 0; i < 64; i++) { e[A.charAt(i)] = i; }
            for (x = 0; x < L; x++) {
                c = e[encodedString.charAt(x)];
                b = (b << 6) + c;
                l += 6;
                while (l >= 8) {
                    ((a = (b >>> (l -= 8)) & 0xff) || (x < (L - 2))) && (r += w(a));
                }
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
