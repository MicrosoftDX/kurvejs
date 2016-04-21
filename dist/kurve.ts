// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module kurve  {
    
// Adapted from the original source: https://github.com/DirtyHairy/typescript-deferred



  export class Error {
        public status: number;
        public statusText: string;
        public text: string;
        public other: any;
    }
    
    function DispatchDeferred(closure: () => void) {
        setTimeout(closure, 0);
    }

     enum PromiseState { Pending, ResolutionInProgress, Resolved, Rejected }

    export interface PromiseCallback<T> {
        (error: Error, result?: T): void;
    }

    class Client {
        constructor(
            private _dispatcher: (closure: () => void) => void,
            private _successCB: any,
            private _errorCB: any
        ) {
            this.result = new Deferred<any, any>(_dispatcher);
        }

        resolve(value: any, defer: boolean): void {
            if (typeof (this._successCB) !== 'function') {
                this.result.resolve(value);
                return;
            }

            if (defer) {
                this._dispatcher(() => this._dispatchCallback(this._successCB, value));
            } else {
                this._dispatchCallback(this._successCB, value);
            }
        }

        reject(error: any, defer: boolean): void {
            if (typeof (this._errorCB) !== 'function') {
                this.result.reject(error);
                return;
            }

            if (defer) {
                this._dispatcher(() => this._dispatchCallback(this._errorCB, error));
            } else {
                this._dispatchCallback(this._errorCB, error);
            }
        }

        private _dispatchCallback(callback: (arg: any) => any, arg: any): void {
            var result: any,
                then: any,
                type: string;

            try {
                result = callback(arg);
                this.result.resolve(result);
            } catch (err) {
                this.result.reject(err);
                return;
            }
        }

        result: Deferred<any, any>;
    }

    export class Deferred<T, E>  {
        private _dispatcher: (closure: () => void)=> void;

        constructor();
        constructor(dispatcher: (closure: () => void) => void);
        constructor(dispatcher?: (closure: () => void) => void) {
            if (dispatcher)
                this._dispatcher = dispatcher;
            else
                this._dispatcher = DispatchDeferred;
            this.promise = new Promise<T, E>(this);
        }

        private DispatchDeferred(closure: () => void) {
            setTimeout(closure, 0);
        }

        then(successCB: any, errorCB: any): any {
            if (typeof (successCB) !== 'function' && typeof (errorCB) !== 'function') {
                return this.promise;
            }

            var client = new Client(this._dispatcher, successCB, errorCB);

            switch (this._state) {
                case PromiseState.Pending:
                case PromiseState.ResolutionInProgress:
                    this._stack.push(client);
                    break;

                case PromiseState.Resolved:
                    client.resolve(this._value, true);
                    break;

                case PromiseState.Rejected:
                    client.reject(this._error, true);
                    break;
            }

            return client.result.promise;
        }

        resolve(value?: T): Deferred<T, E>;

        resolve(value?: Promise<T, E>): Deferred<T, E>;

        resolve(value?: any): Deferred<T, E> {
            if (this._state !== PromiseState.Pending) {
                return this;
            }

            return this._resolve(value);
        }

        private _resolve(value: any): Deferred<T, E> {
            var type = typeof (value),
                then: any,
                pending = true;

            try {
                if (value !== null &&
                    (type === 'object' || type === 'function') &&
                    typeof (then = value.then) === 'function') {
                    if (value === this.promise) {
                        throw new TypeError('recursive resolution');
                    }

                    this._state = PromiseState.ResolutionInProgress;
                    then.call(value,
                        (result: any): void => {
                            if (pending) {
                                pending = false;
                                this._resolve(result);
                            }
                        },
                        (error: any): void => {
                            if (pending) {
                                pending = false;
                                this._reject(error);
                            }
                        }
                    );
                } else {
                    this._state = PromiseState.ResolutionInProgress;

                    this._dispatcher(() => {
                        this._state = PromiseState.Resolved;
                        this._value = value;

                        var i: number,
                            stackSize = this._stack.length;

                        for (i = 0; i < stackSize; i++) {
                            this._stack[i].resolve(value, false);
                        }

                        this._stack.splice(0, stackSize);
                    });
                }
            } catch (err) {
                if (pending) {
                    this._reject(err);
                }
            }

            return this;
        }

        reject(error?: E): Deferred<T, E> {
            if (this._state !== PromiseState.Pending) {
                return this;
            }

            return this._reject(error);
        }

        private _reject(error?: any): Deferred<T, E> {
            this._state = PromiseState.ResolutionInProgress;

            this._dispatcher(() => {
                this._state = PromiseState.Rejected;
                this._error = error;

                var stackSize = this._stack.length,
                    i = 0;

                for (i = 0; i < stackSize; i++) {
                    this._stack[i].reject(error, false);
                }

                this._stack.splice(0, stackSize);
            });

            return this;
        }

        promise: Promise<T, E>;

        private _stack: Array<Client> = [];
        private _state = PromiseState.Pending;
        private _value: T;
        private _error: any;
    }

    export class Promise<T, E>  {
        constructor(private _deferred: Deferred<T, E>) { }

        then<R>(
            successCallback?: (result: T) => R,
            errorCallback?: (error: E) => R
        );

        then(successCallback: any, errorCallback: any): any {
            return this._deferred.then(successCallback, errorCallback);
        }

        fail<R>(
            errorCallback?: (error: E) => R
        );

        fail(errorCallback: any): any {
            return this._deferred.then(undefined, errorCallback);
        }
    }






    export enum EndPointVersion {
        v1=1,
        v2=2
    }

  

    class CachedToken {
        constructor(
            public id: string,
            public scopes: string[],
            public resource: string,
            public token: string,
            public expiry: Date) {};

        public get isExpired() {
            return this.expiry <= new Date(new Date().getTime() + 60000);
        }

        public hasScopes(requiredScopes: string[]) {
            if (!this.scopes) {
                return false;
            }

            return requiredScopes.every(requiredScope => {
                return this.scopes.some(actualScope => requiredScope === actualScope);
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
                tokenStorage.getAll().forEach(({ id, scopes, resource, token, expiry }) => {
                    var cachedToken = new CachedToken(id, scopes, resource, token, new Date(expiry));
                    if (cachedToken.isExpired) {
                        this.tokenStorage.remove(cachedToken.id);
                    } else {
                        this.cachedTokens[cachedToken.id] = cachedToken;
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
        version: EndPointVersion;
        tokenStorage?: TokenStorage;
    }

    export class Identity {
//      public authContext: any = null;
//      public config: any = null;
//      public isCallback: boolean = false;
        public clientId: string;
//      private req: XMLHttpRequest;
        private state: string;
        private version: EndPointVersion;
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
                this.version = EndPointVersion.v1;

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

            var token = new CachedToken(key, scopes, resource, accessToken, expiryDate);
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

        public getCurrentEndPointVersion(): EndPointVersion {
            return this.version;
        }

        public getAccessTokenAsync(resource: string): Promise<string,Error> {

            var d = new Deferred<string,Error>();
            this.getAccessToken(resource, ((error, token) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        }

        public getAccessToken(resource: string, callback: PromiseCallback<string>): void {
            if (this.version !== EndPointVersion.v1) {
                var e = new Error();
                e.statusText = "Currently this identity class is using v2 OAuth mode. You need to use getAccessTokenForScopes() method";
                callback(e);
                return;
            }

            var token = this.tokenCache.getForResource(resource);
            if (token) {
                return callback(null, token.token);
            }

            //If we got this far, we need to go get this token

            //Need to create the iFrame to invoke the acquire token
            this.getTokenCallback = ((token: string, error: Error) => {
                if (error) {
                    callback(error);
                }
                else {
                    this.decodeAccessToken(token, resource);
                    callback(null, token);
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

        public getAccessTokenForScopes(scopes: string[], promptForConsent, callback: (token: string, error: Error) => void): void {
            if (this.version !== EndPointVersion.v2) {
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

            if (loginSettings.scopes && this.version === EndPointVersion.v1) {
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
            if (this.version === EndPointVersion.v2) {
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




    export class Graph {
        private req: XMLHttpRequest = null;
        private accessToken: string = null;
        KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/v1.0";

        constructor(identityInfo: { identity: Identity });
        constructor(identityInfo: { defaultAccessToken: string });
        constructor(identityInfo: any) {
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }

        get me() { return new User(this, this.baseUrl); }
        get users() { return new Users(this, this.baseUrl); }

        public Get<Model, N extends Node>(path:string, self:N, scopes?:string[]): Promise<Singleton<Model, N>, Error> {
            console.log("GET", path, scopes);
            var d = new Deferred<Singleton<Model, N>, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Singleton<Model, N>(jsonResult, self));
            }, null, scopes);

            return d.promise;
         }

        public GetCollection<Model, N extends CollectionNode>(path:string, self:N, next:N, scopes?:string[]): Promise<Collection<Model, N>, Error> {
            console.log("GET collection", path, scopes);
            var d = new Deferred<Collection<Model, N>, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Collection<Model,N>(jsonResult, self, next))
            }, null, scopes);

            return d.promise;
         }

        public Post<Model, N extends Node>(object:Model, path:string, self:N, scopes?:string[]): Promise<Singleton<Model, N>, Error> {
            console.log("POST", path, scopes);
            var d = new Deferred<Singleton<Model, N>, Error>();
            
/*
            this.post(object, path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Response<Model, N>({}, self));
            });
*/
            return d.promise;
         }
 
        public get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?:string[]): void {
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4)
                    if (xhr.status === 200)
                        callback(null, responseType ? xhr.response : xhr.responseText);
                    else
                        callback(this.generateError(xhr));
            }

            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(addTokenError);
                }
            }, scopes);
        }

        public post(object:string, url: string, callback: PromiseCallback<string>, responseType?: string, scopes?:string[]): void {
/*
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4)
                    if (xhr.status === 202)
                        callback(null, responseType ? xhr.response : xhr.responseText);
                    else
                        callback(this.generateError(xhr));
            }
            xhr.send(object)
            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(addTokenError);
                }
            }, scopes);
*/
        }

        private generateError(xhr: XMLHttpRequest): Error {
            var response = new Error();
            response.status = xhr.status;
            response.statusText = xhr.statusText;
            if (xhr.responseType === '' || xhr.responseType === 'text')
                response.text = xhr.responseText;
            else
                response.other = xhr.response;
            return response;

        }

        private addAccessTokenAndSend(xhr: XMLHttpRequest, callback: (error: Error) => void, scopes?:string[]): void {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            } else {
                //Using the integrated Identity object

                if (scopes) {
                    //v2 scope based tokens
                    this.KurveIdentity.getAccessTokenForScopes(scopes,false, ((token: string, error: Error) => {
                        if (error)
                            callback(error);
                        else {
                            xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                            xhr.send();
                            callback(null);
                        }
                    }));

                }
                else {
                    //v1 resource based tokens
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, ((error: Error, token: string) => {
                        if (error)
                            callback(error);
                        else {
                            xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                            xhr.send();
                            callback(null);
                        }
                    }));
                }
            }
        }

    }





export interface ItemBody {
    contentType?: string;
    content?: string;
}

export interface EmailAddress {
    name?: string;
    address?: string;
}

export interface Recipient {
    emailAddress?: EmailAddress;
}

export interface UserDataModel {
    businessPhones?: string;
    displayName?: string;
    givenName?: string;
    jobTitle?: string;
    mail?: string;
    mobilePhone?: string;
    officeLocation?: string;
    preferredLanguage?: string;
    surname?: string;
    userPrincipalName?: string;
    id?: string;
}

export interface ProfilePhotoDataModel {
    id?: string;
    height?: Number;
    width?: Number;
}

export interface MessageDataModel {
    attachments?: AttachmentDataModel[];
    bccRecipients?: Recipient[];
    body?: ItemBody;
    bodyPreview?: string;
    categories?: string[]
    ccRecipients?: Recipient[];
    changeKey?: string;
    conversationId?: string;
    createdDateTime?: string;
    from?: Recipient;
    graph?: any;
    hasAttachments?: boolean;
    id?: string;
    importance?: string;
    isDeliveryReceiptRequested?: boolean;
    isDraft?: boolean;
    isRead?: boolean;
    isReadReceiptRequested?: boolean;
    lastModifiedDateTime?: string;
    parentFolderId?: string;
    receivedDateTime?: string;
    replyTo?: Recipient[]
    sender?: Recipient;
    sentDateTime?: string;
    subject?: string;
    toRecipients?: Recipient[];
    webLink?: string;
}

export interface Attendee {
    status?: ResponseStatus;
    type?: string;
    emailAddress?: EmailAddress;
}

export interface DateTimeTimeZone {
    dateTime?: string;
    timeZone?: string;
}

export interface PatternedRecurrence { }

export interface ResponseStatus {
    response?: string;
    time?: string
}

export interface Location {
    displayName?: string;
    address?: any;
}

export interface EventDataModel {
    attendees?: Attendee[];
    body?: ItemBody;
    bodyPreview?: string;
    categories?: string[];
    changeKey?: string;
    createdDateTime?: string;
    end?: DateTimeTimeZone;
    hasAttachments?: boolean;
    iCalUId?: string;
    id?: string;
    IDBCursor?: string;
    importance?: string;
    isAllDay?: boolean;
    isCancelled?: boolean;
    isOrganizer?: boolean;
    isReminderOn?: boolean;
    lastModifiedDateTime?: string;
    location?: Location;
    organizer?: Recipient;
    originalEndTimeZone?: string;
    originalStartTimeZone?: string;
    recurrence?: PatternedRecurrence;
    reminderMinutesBeforeStart?: number;
    responseRequested?: boolean;
    responseStatus?: ResponseStatus;
    sensitivity?: string;
    seriesMasterId?: string;
    showAs?: string;
    start?: DateTimeTimeZone;
    subject?: string;
    type?: string;
    webLink?: string;
}

export interface GroupDataModel {
    id?: string;
    description?: string;
    displayName?: string;
    groupTypes?: string[];
    mail?: string;
    mailEnabled?: Boolean;
    mailNickname?: string;
    onPremisesLastSyncDateTime?: Date;
    onPremisesSecurityIdentifier?: string;
    onPremisesSyncEnabled?: Boolean;
    proxyAddresses?: string[];
    securityEnabled?: Boolean;
    visibility?: string;
}

export interface MailFolderDataModel {
    id?: string;
    displayName?: string;
    childFolderCount?: number;
    unreadItemCount?: number;
    totalItemCount?: number;
}

export interface AttachmentDataModel {
    contentId?: string;
    id?: string;
    isInline?: boolean;
    lastModifiedDateTime?: Date;
    name?: string;
    size?: number;

    /* File Attachments */
    contentBytes?: string;
    contentLocation?: string;
    contentType?: string;
}




/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing and see how intellisense helps you explore the graph:

    graph.                      me, users
    graph.me.                   events, messages, calendarView, mailFolders, GetUser, odata, select, ...
    graph.me.events.            GetEvents, $, odata, select, ...
    graph.me.events.$           (eventId:string) => Event
    graph.me.events.$("123").   GetEvent, odata, select, ...

Each endpoint exposes the set of available Graph operations through strongly typed methods:

    graph.me.GetUser() => UserDataModel
        GET "/me"
    graph.me.events.GetEvents => EventDataModel[]
        GET "/me/events"
    graph.me.events.CreateEvent(event:EventDataModel) 
        POST "/me/events"

Certain Graph endpoints are implemented as OData "Functions". These are not treated as Graph nodes. They're just methods: 

    graph.me.events.$("123").DeclineEvent(eventResponse:EventResponse)
        POST "/me/events/123/microsoft.graph.decline

Graph operations are exposed through Promises:

    graph.me.messages
    .GetMessages()
    .then(collection =>
        collection.items.forEach(message =>
            console.log(message.subject)
        )
    )

All operations return a "self" property which allows you to continue along the Graph path from the point where you left off:

    graph.me.messages.$("123").GetMessage().then(singleton =>
        console.log(singleton.item.subject);
        singleton.self.attachments.GetAttachments().then(collection => // singleton.self === graph.me.messages.$("123")
            collection.items.forEach(attachment => 
                console.log(attachment.contentBytes)
            )
        )
    )

Operations which return paginated collections can return a "next" request object. This can be utilized in a recursive function:

    ListMessageSubjects(messages:Messages) {
        messages.GetMessages().then(collection => {
            collection.items.forEach(message => console.log(message.subject));
            if (collection.next)
                ListMessageSubjects(collection.next);
        })
    }
    ListMessageSubjects(graph.me.messages);

Every Graph operation may include OData queries:

    graph.me.messages.GetMessages("$select=subject,id&$orderby=id")
        /me/messages/$select=subject,id&$orderby=id

There is an optional OData helper to aid in constructing more complex queries:

    graph.me.messages.GetMessages(new OData()
        .select("subject", "id")
        .orderby("id")
    )
        /me/messages/$select=subject,id&$orderby=id

Some operations include parameters which transform into OData queries 

    graph.me.calendarView.GetCalendarView([start],[end], [odataQuery])
        /me/calendarView?startDateTime=[start]&endDateTime=[end]&[odataQuery]

Note: This initial stab only includes a few familiar pieces of the Microsoft Graph.
*/


export class Scopes {
    private static rootUrl = "https://graph.microsoft.com/";
    static General = {
        OpenId: "openid",
        OfflineAccess: "offline_access",
    }
    static User = {
        Read: Scopes.rootUrl + "User.Read",
        ReadAll: Scopes.rootUrl + "User.Read.All",
        ReadWrite: Scopes.rootUrl + "User.ReadWrite",
        ReadWriteAll: Scopes.rootUrl + "User.ReadWrite.All",
        ReadBasicAll: Scopes.rootUrl + "User.ReadBasic.All",
    }
    static Contacts = {
        Read: Scopes.rootUrl + "Contacts.Read",
        ReadWrite: Scopes.rootUrl + "Contacts.ReadWrite",
    }
    static Directory = {
        ReadAll: Scopes.rootUrl + "Directory.Read.All",
        ReadWriteAll: Scopes.rootUrl + "Directory.ReadWrite.All",
        AccessAsUserAll: Scopes.rootUrl + "Directory.AccessAsUser.All",
    }
    static Group = {
        ReadAll: Scopes.rootUrl + "Group.Read.All",
        ReadWriteAll: Scopes.rootUrl + "Group.ReadWrite.All",
        AccessAsUserAll: Scopes.rootUrl + "Directory.AccessAsUser.All"
    }
    static Mail = {
        Read: Scopes.rootUrl + "Mail.Read",
        ReadWrite: Scopes.rootUrl + "Mail.ReadWrite",
        Send: Scopes.rootUrl + "Mail.Send",
    }
    static Calendars = {
        Read: Scopes.rootUrl + "Calendars.Read",
        ReadWrite: Scopes.rootUrl + "Calendars.ReadWrite",
    }
    static Files = {
        Read: Scopes.rootUrl + "Files.Read",
        ReadAll: Scopes.rootUrl + "Files.Read.All",
        ReadWrite: Scopes.rootUrl + "Files.ReadWrite",
        ReadWriteAppFolder: Scopes.rootUrl + "Files.ReadWrite.AppFolder",
        ReadWriteSelected: Scopes.rootUrl + "Files.ReadWrite.Selected",
    }
    static Tasks = {
        ReadWrite: Scopes.rootUrl + "Tasks.ReadWrite",
    }
    static People = {
        Read: Scopes.rootUrl + "People.Read",
        ReadWrite: Scopes.rootUrl + "People.ReadWrite",
    }
    static Notes = {
        Create: Scopes.rootUrl + "Notes.Create",
        ReadWriteCreatedByApp: Scopes.rootUrl + "Notes.ReadWrite.CreatedByApp",
        Read: Scopes.rootUrl + "Notes.Read",
        ReadAll: Scopes.rootUrl + "Notes.Read.All",
        ReadWriteAll: Scopes.rootUrl + "Notes.ReadWrite.All",
    }
}

let queryUnion = (query1:string, query2:string) => (query1 ? query1 + (query2 ? "&" + query2 : "" ) : query2); 

export class OData {
    constructor(public query?:string) {
    }
    
    toString = () => this.query;

    odata = (query:string) => {
        this.query = queryUnion(this.query, query);
        return this;
    }

    select   = (...fields:string[])  => this.odata(`$select=${fields.join(",")}`);
    expand   = (...fields:string[])  => this.odata(`$expand=${fields.join(",")}`);
    filter   = (query:string)        => this.odata(`$filter=${query}`);
    orderby  = (...fields:string[])  => this.odata(`$orderby=${fields.join(",")}`);
    top      = (items:Number)        => this.odata(`$top=${items}`);
    skip     = (items:Number)        => this.odata(`$skip=${items}`);
}

type ODataQuery = OData | string;

let pathWithQuery = (path:string, odataQuery?:ODataQuery) => {
    let query = odataQuery && odataQuery.toString();
    return path + (query ? "?" + query : "");
}

export class Singleton<Model, N extends Node> {
    constructor(public raw:any, public self:N) {
    }

    get item() {
        return this.raw as Model;
    }
}

export class Collection<Model, N extends CollectionNode> {
    constructor(public raw:any, public self:N, public next:N) {
        let nextLink = this.raw["@odata.nextLink"];
        if (nextLink) {
            this.next.nextLink = nextLink;
        } else {
            this.next = undefined;
        }
    }

    get items() {
        return (this.raw.value ? this.raw.value : [this.raw]) as Model[];
    }
}

export abstract class Node {
    constructor(protected graph:Graph, protected path:string) {
    }

    //Only adds scopes when linked to a v2 Oauth of kurve identity
    protected scopesForV2 = (scopes: string[]) =>
        this.graph.KurveIdentity && this.graph.KurveIdentity.getCurrentEndPointVersion() === EndPointVersion.v2 ? scopes : null;
    
    pathWithQuery = (odataQuery?:ODataQuery, pathSuffix:string = "") => pathWithQuery(this.path + pathSuffix, odataQuery);
}

export abstract class CollectionNode extends Node {    
    private _nextLink:string;   // this is only set when the collection in question is from a nextLink

    pathWithQuery = (odataQuery?:ODataQuery, pathSuffix:string = "") => this._nextLink || pathWithQuery(this.path + pathSuffix, odataQuery);
    
    set nextLink(pathWithQuery:string) {
        this._nextLink = pathWithQuery;
    }
}

export class Attachment extends Node {
    constructor(graph:Graph, path:string="", private context:string, attachmentId?:string) {
        super(graph, path + (attachmentId ? "/" + attachmentId : ""));
    }

    static scopes = {
        messages: [Scopes.Mail.Read],
        events: [Scopes.Calendars.Read],
    }

    GetAttachment = (odataQuery?:ODataQuery) => this.graph.Get<AttachmentDataModel, Attachment>(this.pathWithQuery(odataQuery), this, this.scopesForV2(Attachment.scopes[this.context]));
/*    
    PATCH = this.graph.PATCH<AttachmentDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<AttachmentDataModel>(this.path, this.query);
*/
}

export class Attachments extends CollectionNode {
    constructor(graph:Graph, path:string="", private context:string) {
        super(graph, path + "/attachments");
    }

    $ = (attachmentId:string) => new Attachment(this.graph, this.path, this.context, attachmentId);
    
    GetAttachments = (odataQuery?:ODataQuery) => this.graph.GetCollection<AttachmentDataModel, Attachments>(this.pathWithQuery(odataQuery), this, new Attachments(this.graph, null, this.context), this.scopesForV2(Attachment.scopes[this.context]));
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

export class Message extends Node {
    constructor(graph:Graph, path:string="", messageId?:string) {
        super(graph, path + (messageId ? "/" + messageId : ""));
    }
    
    get attachments() { return new Attachments(this.graph, this.path, "messages"); }

    GetMessage  = (odataQuery?:ODataQuery) => this.graph.Get<MessageDataModel, Message>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.Read]));
    SendMessage = (odataQuery?:ODataQuery) => this.graph.Post<MessageDataModel, Message>(null, this.pathWithQuery(odataQuery, "/microsoft.graph.sendMail"), this, this.scopesForV2([Scopes.Mail.Send]));
/*
    PATCH = this.graph.PATCH<MessageDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<MessageDataModel>(this.path, this.query);
*/
}

export class Messages extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/messages");
    }

    $ = (messageId:string) => new Message(this.graph, this.path, messageId);

    GetMessages     = (odataQuery?:ODataQuery) => this.graph.GetCollection<MessageDataModel, Messages>(this.pathWithQuery(odataQuery), this, new Messages(this.graph), this.scopesForV2([Scopes.Mail.Read]));
    CreateMessage   = (object:MessageDataModel, odataQuery?:ODataQuery) => this.graph.Post<MessageDataModel, Messages>(object, this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.ReadWrite]));
}

export class Event extends Node {
    constructor(graph:Graph, path:string="", eventId:string) {
        super(graph, path + (eventId ? "/" + eventId : ""));
    }

    get attachments() { return new Attachments(this.graph, this.path, "events"); }

    GetEvent = (odataQuery?:ODataQuery) => this.graph.Get<EventDataModel, Event>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Calendars.Read]));
/*
    PATCH = this.graph.PATCH<EventDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<EventDataModel>(this.path, this.query);
*/
}


export class Events extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/events");
    }

    $ = (eventId:string) => new Event(this.graph, this.path, eventId);

    GetEvents = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, Events>(this.pathWithQuery(odataQuery), this, new Events(this.graph), this.scopesForV2([Scopes.Calendars.Read]));
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

export class CalendarView extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/calendarView");
    }
    
    dateRange = (startDate:Date, endDate:Date) => `startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`

    GetCalendarView = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, CalendarView>(this.pathWithQuery(odataQuery), this, new CalendarView(this.graph), this.scopesForV2([Scopes.Calendars.Read]));
}


export class MailFolder extends Node {
    constructor(graph:Graph, path:string="", mailFolderId:string) {
        super(graph, path + (mailFolderId ? "/" + mailFolderId : ""));
    }


    GetMailFolder = (odataQuery?:ODataQuery) => this.graph.Get<MailFolderDataModel, MailFolder>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.Read]));
}

export class MailFolders extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/mailFolders");
    }

    $ = (mailFolderId:string) => new MailFolder(this.graph, this.path, mailFolderId);

    GetMailFolders = (odataQuery?:ODataQuery) => this.graph.GetCollection<MailFolderDataModel, MailFolders>(this.pathWithQuery(odataQuery), this, new MailFolders(this.graph), this.scopesForV2([Scopes.Mail.Read]));
}


export class Photo extends Node {    
    constructor(graph:Graph, path:string="", private context:string) {
        super(graph, path + "/photo" );
    }

    static scopes = {
        user: [Scopes.User.ReadBasicAll],
        group: [Scopes.Group.ReadAll],
        contact: [Scopes.Contacts.Read]
    }

    GetPhotoProperties = (odataQuery?:ODataQuery) => this.graph.Get<ProfilePhotoDataModel, Photo>(this.pathWithQuery(odataQuery), this, this.scopesForV2(Photo.scopes[this.context]));
    GetPhotoImage = (odataQuery?:ODataQuery) => this.graph.Get<any, Photo>(this.pathWithQuery(odataQuery, "/$value"), this, this.scopesForV2(Photo.scopes[this.context]));
}

export class Manager extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/manager" );
    }

    GetManager = (odataQuery?:ODataQuery) => this.graph.Get<UserDataModel, Manager>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.User.ReadAll]));
}

export class MemberOf extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/memberOf");
    }

    GetGroups = (odataQuery?:ODataQuery) => this.graph.GetCollection<GroupDataModel, MemberOf>(this.pathWithQuery(odataQuery), this, new MemberOf(this.graph), this.scopesForV2([Scopes.User.ReadAll]));
}

export class User extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, userId ? path + "/" + userId : path + "/me");
    }

    get messages()      { return new Messages(this.graph, this.path); }
    get events()        { return new Events(this.graph, this.path); }
    get calendarView()  { return new CalendarView(this.graph, this.path); }
    get mailFolders()   { return new MailFolders(this.graph, this.path) }
    get photo()         { return new Photo(this.graph, this.path, "user"); }
    get manager()       { return new Manager(this.graph, this.path); }

    GetUser = (odataQuery?:ODataQuery) => this.graph.Get<UserDataModel, User>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.User.Read]));
/*
    PATCH = this.graph.PATCH<UserDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<UserDataModel>(this.path, this.query);
*/
}

export class Users extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/users");
    }

    $ = (userId:string) => new User(this.graph, this.path, userId);

    GetUsers = (odataQuery?:ODataQuery) => this.graph.GetCollection<UserDataModel, Users>(this.pathWithQuery(odataQuery), this, new Users(this.graph), this.scopesForV2([Scopes.User.Read]));
/*
    CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
*/
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

