var __extends = (this && this.__extends) || function (d, b) {
for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
function __() { this.constructor = d; }
__.prototype = b.prototype;
d.prototype = new __();
};
// Adapted from the original source: https://github.com/DirtyHairy/typescript-deferred
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Kurve;
(function (Kurve) {
    function DispatchDeferred(closure) {
        setTimeout(closure, 0);
    }
    var PromiseState;
    (function (PromiseState) {
        PromiseState[PromiseState["Pending"] = 0] = "Pending";
        PromiseState[PromiseState["ResolutionInProgress"] = 1] = "ResolutionInProgress";
        PromiseState[PromiseState["Resolved"] = 2] = "Resolved";
        PromiseState[PromiseState["Rejected"] = 3] = "Rejected";
    })(PromiseState || (PromiseState = {}));
    var Client = (function () {
        function Client(_dispatcher, _successCB, _errorCB) {
            this._dispatcher = _dispatcher;
            this._successCB = _successCB;
            this._errorCB = _errorCB;
            this.result = new Deferred(_dispatcher);
        }
        Client.prototype.resolve = function (value, defer) {
            var _this = this;
            if (typeof (this._successCB) !== 'function') {
                this.result.resolve(value);
                return;
            }
            if (defer) {
                this._dispatcher(function () { return _this._dispatchCallback(_this._successCB, value); });
            }
            else {
                this._dispatchCallback(this._successCB, value);
            }
        };
        Client.prototype.reject = function (error, defer) {
            var _this = this;
            if (typeof (this._errorCB) !== 'function') {
                this.result.reject(error);
                return;
            }
            if (defer) {
                this._dispatcher(function () { return _this._dispatchCallback(_this._errorCB, error); });
            }
            else {
                this._dispatchCallback(this._errorCB, error);
            }
        };
        Client.prototype._dispatchCallback = function (callback, arg) {
            var result, then, type;
            try {
                result = callback(arg);
                this.result.resolve(result);
            }
            catch (err) {
                this.result.reject(err);
                return;
            }
        };
        return Client;
    })();
    var Deferred = (function () {
        function Deferred(dispatcher) {
            this._stack = [];
            this._state = PromiseState.Pending;
            if (dispatcher)
                this._dispatcher = dispatcher;
            else
                this._dispatcher = DispatchDeferred;
            this.promise = new Promise(this);
        }
        Deferred.prototype.DispatchDeferred = function (closure) {
            setTimeout(closure, 0);
        };
        Deferred.prototype.then = function (successCB, errorCB) {
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
        };
        Deferred.prototype.resolve = function (value) {
            if (this._state !== PromiseState.Pending) {
                return this;
            }
            return this._resolve(value);
        };
        Deferred.prototype._resolve = function (value) {
            var _this = this;
            var type = typeof (value), then, pending = true;
            try {
                if (value !== null &&
                    (type === 'object' || type === 'function') &&
                    typeof (then = value.then) === 'function') {
                    if (value === this.promise) {
                        throw new TypeError('recursive resolution');
                    }
                    this._state = PromiseState.ResolutionInProgress;
                    then.call(value, function (result) {
                        if (pending) {
                            pending = false;
                            _this._resolve(result);
                        }
                    }, function (error) {
                        if (pending) {
                            pending = false;
                            _this._reject(error);
                        }
                    });
                }
                else {
                    this._state = PromiseState.ResolutionInProgress;
                    this._dispatcher(function () {
                        _this._state = PromiseState.Resolved;
                        _this._value = value;
                        var i, stackSize = _this._stack.length;
                        for (i = 0; i < stackSize; i++) {
                            _this._stack[i].resolve(value, false);
                        }
                        _this._stack.splice(0, stackSize);
                    });
                }
            }
            catch (err) {
                if (pending) {
                    this._reject(err);
                }
            }
            return this;
        };
        Deferred.prototype.reject = function (error) {
            if (this._state !== PromiseState.Pending) {
                return this;
            }
            return this._reject(error);
        };
        Deferred.prototype._reject = function (error) {
            var _this = this;
            this._state = PromiseState.ResolutionInProgress;
            this._dispatcher(function () {
                _this._state = PromiseState.Rejected;
                _this._error = error;
                var stackSize = _this._stack.length, i = 0;
                for (i = 0; i < stackSize; i++) {
                    _this._stack[i].reject(error, false);
                }
                _this._stack.splice(0, stackSize);
            });
            return this;
        };
        return Deferred;
    })();
    Kurve.Deferred = Deferred;
    var Promise = (function () {
        function Promise(_deferred) {
            this._deferred = _deferred;
        }
        Promise.prototype.then = function (successCallback, errorCallback) {
            return this._deferred.then(successCallback, errorCallback);
        };
        Promise.prototype.fail = function (errorCallback) {
            return this._deferred.then(undefined, errorCallback);
        };
        return Promise;
    })();
    Kurve.Promise = Promise;
})(Kurve || (Kurve = {}));
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
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Kurve;
(function (Kurve) {
    (function (OAuthVersion) {
        OAuthVersion[OAuthVersion["v1"] = 1] = "v1";
        OAuthVersion[OAuthVersion["v2"] = 2] = "v2";
    })(Kurve.OAuthVersion || (Kurve.OAuthVersion = {}));
    var OAuthVersion = Kurve.OAuthVersion;
    var Error = (function () {
        function Error() {
        }
        return Error;
    })();
    Kurve.Error = Error;
    var CachedToken = (function () {
        function CachedToken(id, scopes, resource, token, expiry) {
            this.id = id;
            this.scopes = scopes;
            this.resource = resource;
            this.token = token;
            this.expiry = expiry;
        }
        ;
        Object.defineProperty(CachedToken.prototype, "isExpired", {
            get: function () {
                return this.expiry <= new Date(new Date().getTime() + 60000);
            },
            enumerable: true,
            configurable: true
        });
        CachedToken.prototype.hasScopes = function (requiredScopes) {
            var _this = this;
            if (!this.scopes) {
                return false;
            }
            return requiredScopes.every(function (requiredScope) {
                return _this.scopes.some(function (actualScope) { return requiredScope === actualScope; });
            });
        };
        return CachedToken;
    })();
    var TokenCache = (function () {
        function TokenCache(tokenStorage) {
            var _this = this;
            this.tokenStorage = tokenStorage;
            this.cachedTokens = {};
            if (tokenStorage) {
                tokenStorage.getAll().forEach(function (_a) {
                    var id = _a.id, scopes = _a.scopes, resource = _a.resource, token = _a.token, expiry = _a.expiry;
                    var cachedToken = new CachedToken(id, scopes, resource, token, new Date(expiry));
                    if (cachedToken.isExpired) {
                        _this.tokenStorage.remove(cachedToken.id);
                    }
                    else {
                        _this.cachedTokens[cachedToken.id] = cachedToken;
                    }
                });
            }
        }
        TokenCache.prototype.add = function (token) {
            this.cachedTokens[token.id] = token;
            this.tokenStorage && this.tokenStorage.add(token.id, token);
        };
        TokenCache.prototype.getForResource = function (resource) {
            var cachedToken = this.cachedTokens[resource];
            if (cachedToken && cachedToken.isExpired) {
                this.remove(resource);
                return null;
            }
            return cachedToken;
        };
        TokenCache.prototype.getForScopes = function (scopes) {
            for (var key in this.cachedTokens) {
                var cachedToken = this.cachedTokens[key];
                if (cachedToken.hasScopes(scopes)) {
                    if (cachedToken.isExpired) {
                        this.remove(key);
                    }
                    else {
                        return cachedToken;
                    }
                }
            }
            return null;
        };
        TokenCache.prototype.clear = function () {
            this.cachedTokens = {};
            this.tokenStorage && this.tokenStorage.clear();
        };
        TokenCache.prototype.remove = function (key) {
            this.tokenStorage && this.tokenStorage.remove(key);
            delete this.cachedTokens[key];
        };
        return TokenCache;
    })();
    var IdToken = (function () {
        function IdToken() {
        }
        return IdToken;
    })();
    Kurve.IdToken = IdToken;
    var Identity = (function () {
        //      private tenant: string = "";
        function Identity(identitySettings) {
            var _this = this;
            this.policy = "";
            this.clientId = identitySettings.clientId;
            this.tokenProcessorUrl = identitySettings.tokenProcessingUri;
            //          this.req = new XMLHttpRequest();
            if (identitySettings.version)
                this.version = identitySettings.version;
            else
                this.version = OAuthVersion.v1;
            this.tokenCache = new TokenCache(identitySettings.tokenStorage);
            //Callback handler from other windows
            window.addEventListener("message", function (event) {
                if (event.data.type === "id_token") {
                    if (event.data.error) {
                        var e = new Error();
                        e.text = event.data.error;
                        _this.loginCallback(e);
                    }
                    else {
                        //check for state
                        if (_this.state !== event.data.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            _this.loginCallback(error);
                        }
                        else {
                            _this.decodeIdToken(event.data.token);
                            _this.loginCallback(null);
                        }
                    }
                }
                else if (event.data.type === "access_token") {
                    if (event.data.error) {
                        var e = new Error();
                        e.text = event.data.error;
                        _this.getTokenCallback(null, e);
                    }
                    else {
                        var token = event.data.token;
                        var iframe = document.getElementById("tokenIFrame");
                        iframe.parentNode.removeChild(iframe);
                        if (event.data.state !== _this.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            _this.getTokenCallback(null, error);
                        }
                        else {
                            _this.getTokenCallback(token, null);
                        }
                    }
                }
            });
        }
        Identity.prototype.checkForIdentityRedirect = function () {
            function token(s) {
                var start = window.location.href.indexOf(s);
                if (start < 0)
                    return null;
                var end = window.location.href.indexOf("&", start + s.length);
                return window.location.href.substring(start, ((end > 0) ? end : window.location.href.length));
            }
            function parseQueryString(str) {
                var queryString = str || window.location.search || '';
                var keyValPairs = [];
                var params = {};
                queryString = queryString.replace(/.*?\?/, "");
                if (queryString.length) {
                    keyValPairs = queryString.split('&');
                    for (var pairNum in keyValPairs) {
                        var key = keyValPairs[pairNum].split('=')[0];
                        if (!key.length)
                            continue;
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
                if (true || this.state === params["state"][0]) {
                    this.decodeIdToken(idToken);
                    this.loginCallback && this.loginCallback(null);
                }
                else {
                    var error = new Error();
                    error.statusText = "Invalid state";
                    this.loginCallback && this.loginCallback(error);
                }
                return true;
            }
            else if (accessToken) {
                throw "Should not get here.  This should be handled via the iframe approach.";
            }
            return false;
        };
        Identity.prototype.decodeIdToken = function (idToken) {
            var _this = this;
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
            var expiration = expiryDate.getTime() - new Date().getTime() - 300000;
            this.refreshTimer = setTimeout((function () {
                _this.renewIdToken();
            }), expiration);
        };
        Identity.prototype.decodeAccessToken = function (accessToken, resource, scopes) {
            var decodedToken = this.base64Decode(accessToken.substring(accessToken.indexOf('.') + 1, accessToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            var key = resource || scopes.join(" ");
            var token = new CachedToken(key, scopes, resource, accessToken, expiryDate);
            this.tokenCache.add(token);
        };
        Identity.prototype.getIdToken = function () {
            return this.idToken;
        };
        Identity.prototype.isLoggedIn = function () {
            if (!this.idToken)
                return false;
            return (this.idToken.Expiry > new Date());
        };
        Identity.prototype.renewIdToken = function () {
            clearTimeout(this.refreshTimer);
            this.login(function () { });
        };
        Identity.prototype.getCurrentOauthVersion = function () {
            return this.version;
        };
        Identity.prototype.getAccessTokenAsync = function (resource) {
            var d = new Kurve.Deferred();
            this.getAccessToken(resource, (function (token, error) {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        };
        Identity.prototype.getAccessToken = function (resource, callback) {
            var _this = this;
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
            this.getTokenCallback = (function (token, error) {
                if (error) {
                    callback(null, error);
                }
                else {
                    _this.decodeAccessToken(token, resource);
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
        };
        Identity.prototype.getAccessTokenForScopesAsync = function (scopes, promptForConsent) {
            if (promptForConsent === void 0) { promptForConsent = false; }
            var d = new Kurve.Deferred();
            this.getAccessTokenForScopes(scopes, promptForConsent, function (token, error) {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(token);
                }
            });
            return d.promise;
        };
        Identity.prototype.getAccessTokenForScopes = function (scopes, promptForConsent, callback) {
            var _this = this;
            if (promptForConsent === void 0) { promptForConsent = false; }
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
            this.getTokenCallback = (function (token, error) {
                if (error) {
                    if (promptForConsent || !error.text) {
                        callback(null, error);
                    }
                    else if (error.text.indexOf("AADSTS65001") >= 0) {
                        //We will need to try getting the consent
                        _this.getAccessTokenForScopes(scopes, true, _this.getTokenCallback);
                    }
                    else {
                        callback(null, error);
                    }
                }
                else {
                    _this.decodeAccessToken(token, null, scopes);
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
            }
            else {
                window.open(this.tokenProcessorUrl + "?clientId=" + encodeURIComponent(this.clientId) +
                    "&scopes=" + encodeURIComponent(scopes.join(" ")) +
                    "&redirectUri=" + encodeURIComponent(this.tokenProcessorUrl) +
                    "&version=" + encodeURIComponent(this.version.toString()) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce) +
                    "&op=token", "_blank");
            }
        };
        Identity.prototype.loginAsync = function (loginSettings) {
            var d = new Kurve.Deferred();
            this.login(function (error) {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }, loginSettings);
            return d.promise;
        };
        Identity.prototype.login = function (callback, loginSettings) {
            this.loginCallback = callback;
            if (!loginSettings)
                loginSettings = {};
            if (loginSettings.policy)
                this.policy = loginSettings.policy;
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
                if (!loginSettings.scopes)
                    loginSettings.scopes = [];
                if (loginSettings.scopes.indexOf("profile") < 0)
                    loginSettings.scopes.push("profile");
                if (loginSettings.scopes.indexOf("openid") < 0)
                    loginSettings.scopes.push("openid");
                loginURL += "&scopes=" + encodeURIComponent(loginSettings.scopes.join(" "));
            }
            window.open(loginURL, "_blank");
        };
        Identity.prototype.loginNoWindowAsync = function (toUrl) {
            var d = new Kurve.Deferred();
            this.loginNoWindow(function (error) {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }, toUrl);
            return d.promise;
        };
        Identity.prototype.loginNoWindow = function (callback, toUrl) {
            this.loginCallback = callback;
            this.state = "clientId=" + this.clientId + "&" + "tokenProcessorUrl=" + this.tokenProcessorUrl;
            this.nonce = this.generateNonce();
            var redirected = this.checkForIdentityRedirect();
            if (!redirected) {
                var redirectUri = (toUrl) ? toUrl : window.location.href.split("#")[0]; // default the no login window scenario to return to the current page
                var url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token" +
                    "&client_id=" + encodeURIComponent(this.clientId) +
                    "&redirect_uri=" + encodeURIComponent(redirectUri) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce);
                window.location.href = url;
            }
        };
        Identity.prototype.logOut = function () {
            this.tokenCache.clear();
            var url = "https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=" + encodeURI(window.location.href);
            window.location.href = url;
        };
        Identity.prototype.base64Decode = function (encodedString) {
            var e = {}, i, b = 0, c, x, l = 0, a, r = '', w = String.fromCharCode, L = encodedString.length;
            var A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            for (i = 0; i < 64; i++) {
                e[A.charAt(i)] = i;
            }
            for (x = 0; x < L; x++) {
                c = e[encodedString.charAt(x)];
                b = (b << 6) + c;
                l += 6;
                while (l >= 8) {
                    ((a = (b >>> (l -= 8)) & 0xff) || (x < (L - 2))) && (r += w(a));
                }
            }
            return r;
        };
        Identity.prototype.generateNonce = function () {
            var text = "";
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            for (var i = 0; i < 32; i++) {
                text += chars.charAt(Math.floor(Math.random() * chars.length));
            }
            return text;
        };
        return Identity;
    })();
    Kurve.Identity = Identity;
})(Kurve || (Kurve = {}));
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
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Kurve;
(function (Kurve) {
    var Scopes;
    (function (Scopes) {
        var Util = (function () {
            function Util() {
            }
            Util.rootUrl = "https://graph.microsoft.com/";
            return Util;
        })();
        var General = (function () {
            function General() {
            }
            General.OpenId = "openid";
            General.OfflineAccess = "offline_access";
            return General;
        })();
        Scopes.General = General;
        var User = (function () {
            function User() {
            }
            User.Read = Util.rootUrl + "User.Read";
            User.ReadWrite = Util.rootUrl + "User.ReadWrite";
            User.ReadBasicAll = Util.rootUrl + "User.ReadBasic.All";
            User.ReadAll = Util.rootUrl + "User.Read.All";
            User.ReadWriteAll = Util.rootUrl + "User.ReadWrite.All";
            return User;
        })();
        Scopes.User = User;
        var Contacts = (function () {
            function Contacts() {
            }
            Contacts.Read = Util.rootUrl + "Contacts.Read";
            Contacts.ReadWrite = Util.rootUrl + "Contacts.ReadWrite";
            return Contacts;
        })();
        Scopes.Contacts = Contacts;
        var Directory = (function () {
            function Directory() {
            }
            Directory.ReadAll = Util.rootUrl + "Directory.Read.All";
            Directory.ReadWriteAll = Util.rootUrl + "Directory.ReadWrite.All";
            Directory.AccessAsUserAll = Util.rootUrl + "Directory.AccessAsUser.All";
            return Directory;
        })();
        Scopes.Directory = Directory;
        var Group = (function () {
            function Group() {
            }
            Group.ReadAll = Util.rootUrl + "Group.Read.All";
            Group.ReadWriteAll = Util.rootUrl + "Group.ReadWrite.All";
            Group.AccessAsUserAll = Util.rootUrl + "Directory.AccessAsUser.All";
            return Group;
        })();
        Scopes.Group = Group;
        var Mail = (function () {
            function Mail() {
            }
            Mail.Read = Util.rootUrl + "Mail.Read";
            Mail.ReadWrite = Util.rootUrl + "Mail.ReadWrite";
            Mail.Send = Util.rootUrl + "Mail.Send";
            return Mail;
        })();
        Scopes.Mail = Mail;
        var Calendars = (function () {
            function Calendars() {
            }
            Calendars.Read = Util.rootUrl + "Calendars.Read";
            Calendars.ReadWrite = Util.rootUrl + "Calendars.ReadWrite";
            return Calendars;
        })();
        Scopes.Calendars = Calendars;
        var Files = (function () {
            function Files() {
            }
            Files.Read = Util.rootUrl + "Files.Read";
            Files.ReadAll = Util.rootUrl + "Files.Read.All";
            Files.ReadWrite = Util.rootUrl + "Files.ReadWrite";
            Files.ReadWriteAppFolder = Util.rootUrl + "Files.ReadWrite.AppFolder";
            Files.ReadWriteSelected = Util.rootUrl + "Files.ReadWrite.Selected";
            return Files;
        })();
        Scopes.Files = Files;
        var Tasks = (function () {
            function Tasks() {
            }
            Tasks.ReadWrite = Util.rootUrl + "Tasks.ReadWrite";
            return Tasks;
        })();
        Scopes.Tasks = Tasks;
        var People = (function () {
            function People() {
            }
            People.Read = Util.rootUrl + "People.Read";
            People.ReadWrite = Util.rootUrl + "People.ReadWrite";
            return People;
        })();
        Scopes.People = People;
        var Notes = (function () {
            function Notes() {
            }
            Notes.Create = Util.rootUrl + "Notes.Create";
            Notes.ReadWriteCreatedByApp = Util.rootUrl + "Notes.ReadWrite.CreatedByApp";
            Notes.Read = Util.rootUrl + "Notes.Read";
            Notes.ReadAll = Util.rootUrl + "Notes.Read.All";
            Notes.ReadWriteAll = Util.rootUrl + "Notes.ReadWrite.All";
            return Notes;
        })();
        Scopes.Notes = Notes;
    })(Scopes = Kurve.Scopes || (Kurve.Scopes = {}));
    var DataModelWrapper = (function () {
        function DataModelWrapper(graph, _data) {
            this.graph = graph;
            this._data = _data;
        }
        Object.defineProperty(DataModelWrapper.prototype, "data", {
            get: function () { return this._data; },
            enumerable: true,
            configurable: true
        });
        return DataModelWrapper;
    })();
    Kurve.DataModelWrapper = DataModelWrapper;
    var DataModelListWrapper = (function (_super) {
        __extends(DataModelListWrapper, _super);
        function DataModelListWrapper() {
            _super.apply(this, arguments);
        }
        return DataModelListWrapper;
    })(DataModelWrapper);
    Kurve.DataModelListWrapper = DataModelListWrapper;
    var ProfilePhotoDataModel = (function () {
        function ProfilePhotoDataModel() {
        }
        return ProfilePhotoDataModel;
    })();
    Kurve.ProfilePhotoDataModel = ProfilePhotoDataModel;
    var ProfilePhoto = (function (_super) {
        __extends(ProfilePhoto, _super);
        function ProfilePhoto() {
            _super.apply(this, arguments);
        }
        return ProfilePhoto;
    })(DataModelWrapper);
    Kurve.ProfilePhoto = ProfilePhoto;
    var UserDataModel = (function () {
        function UserDataModel() {
        }
        return UserDataModel;
    })();
    Kurve.UserDataModel = UserDataModel;
    (function (EventsEndpoint) {
        EventsEndpoint[EventsEndpoint["events"] = 0] = "events";
        EventsEndpoint[EventsEndpoint["calendarView"] = 1] = "calendarView";
    })(Kurve.EventsEndpoint || (Kurve.EventsEndpoint = {}));
    var EventsEndpoint = Kurve.EventsEndpoint;
    var User = (function (_super) {
        __extends(User, _super);
        function User() {
            _super.apply(this, arguments);
        }
        // These are all passthroughs to the graph
        User.prototype.events = function (callback, odataQuery) {
            this.graph.eventsForUser(this._data.userPrincipalName, EventsEndpoint.events, callback, odataQuery);
        };
        User.prototype.eventsAsync = function (odataQuery) {
            return this.graph.eventsForUserAsync(this._data.userPrincipalName, EventsEndpoint.events, odataQuery);
        };
        User.prototype.memberOf = function (callback, Error, odataQuery) {
            this.graph.memberOfForUser(this._data.userPrincipalName, callback, odataQuery);
        };
        User.prototype.memberOfAsync = function (odataQuery) {
            return this.graph.memberOfForUserAsync(this._data.userPrincipalName, odataQuery);
        };
        User.prototype.messages = function (callback, odataQuery) {
            this.graph.messagesForUser(this._data.userPrincipalName, callback, odataQuery);
        };
        User.prototype.messagesAsync = function (odataQuery) {
            return this.graph.messagesForUserAsync(this._data.userPrincipalName, odataQuery);
        };
        User.prototype.manager = function (callback, odataQuery) {
            this.graph.managerForUser(this._data.userPrincipalName, callback, odataQuery);
        };
        User.prototype.managerAsync = function (odataQuery) {
            return this.graph.managerForUserAsync(this._data.userPrincipalName, odataQuery);
        };
        User.prototype.profilePhoto = function (callback, odataQuery) {
            this.graph.profilePhotoForUser(this._data.userPrincipalName, callback, odataQuery);
        };
        User.prototype.profilePhotoAsync = function (odataQuery) {
            return this.graph.profilePhotoForUserAsync(this._data.userPrincipalName, odataQuery);
        };
        User.prototype.profilePhotoValue = function (callback, odataQuery) {
            this.graph.profilePhotoValueForUser(this._data.userPrincipalName, callback, odataQuery);
        };
        User.prototype.profilePhotoValueAsync = function (odataQuery) {
            return this.graph.profilePhotoValueForUserAsync(this._data.userPrincipalName, odataQuery);
        };
        User.prototype.calendarView = function (callback, odataQuery) {
            this.graph.eventsForUser(this._data.userPrincipalName, EventsEndpoint.calendarView, callback, odataQuery);
        };
        User.prototype.calendarViewAsync = function (odataQuery) {
            return this.graph.eventsForUserAsync(this._data.userPrincipalName, EventsEndpoint.calendarView, odataQuery);
        };
        User.prototype.mailFolders = function (callback, odataQuery) {
            this.graph.mailFoldersForUser(this._data.userPrincipalName, callback, odataQuery);
        };
        User.prototype.mailFoldersAsync = function (odataQuery) {
            return this.graph.mailFoldersForUserAsync(this._data.userPrincipalName, odataQuery);
        };
        User.prototype.message = function (messageId, callback, odataQuery) {
            this.graph.messageForUser(this._data.userPrincipalName, messageId, callback, odataQuery);
        };
        User.prototype.messageAsync = function (messageId, odataQuery) {
            return this.graph.messageForUserAsync(this._data.userPrincipalName, messageId, odataQuery);
        };
        User.prototype.event = function (eventId, callback, odataQuery) {
            this.graph.eventForUser(this._data.userPrincipalName, eventId, callback, odataQuery);
        };
        User.prototype.eventAsync = function (eventId, odataQuery) {
            return this.graph.eventForUserAsync(this._data.userPrincipalName, eventId, odataQuery);
        };
        User.prototype.messageAttachment = function (messageId, attachmentId, callback, odataQuery) {
            this.graph.messageAttachmentForUser(this._data.userPrincipalName, messageId, attachmentId, callback, odataQuery);
        };
        User.prototype.messageAttachmentAsync = function (messageId, attachmentId, odataQuery) {
            return this.graph.messageAttachmentForUserAsync(this._data.userPrincipalName, messageId, attachmentId, odataQuery);
        };
        return User;
    })(DataModelWrapper);
    Kurve.User = User;
    var Users = (function (_super) {
        __extends(Users, _super);
        function Users() {
            _super.apply(this, arguments);
        }
        return Users;
    })(DataModelListWrapper);
    Kurve.Users = Users;
    var MessageDataModel = (function () {
        function MessageDataModel() {
        }
        return MessageDataModel;
    })();
    Kurve.MessageDataModel = MessageDataModel;
    var Message = (function (_super) {
        __extends(Message, _super);
        function Message() {
            _super.apply(this, arguments);
        }
        return Message;
    })(DataModelWrapper);
    Kurve.Message = Message;
    var Messages = (function (_super) {
        __extends(Messages, _super);
        function Messages() {
            _super.apply(this, arguments);
        }
        return Messages;
    })(DataModelListWrapper);
    Kurve.Messages = Messages;
    var EventDataModel = (function () {
        function EventDataModel() {
        }
        return EventDataModel;
    })();
    Kurve.EventDataModel = EventDataModel;
    var Event = (function (_super) {
        __extends(Event, _super);
        function Event() {
            _super.apply(this, arguments);
        }
        return Event;
    })(DataModelWrapper);
    Kurve.Event = Event;
    var Events = (function (_super) {
        __extends(Events, _super);
        function Events(graph, endpoint, _data) {
            _super.call(this, graph, _data);
            this.graph = graph;
            this.endpoint = endpoint;
            this._data = _data;
        }
        return Events;
    })(DataModelListWrapper);
    Kurve.Events = Events;
    var Contact = (function () {
        function Contact() {
        }
        return Contact;
    })();
    Kurve.Contact = Contact;
    var GroupDataModel = (function () {
        function GroupDataModel() {
        }
        return GroupDataModel;
    })();
    Kurve.GroupDataModel = GroupDataModel;
    var Group = (function (_super) {
        __extends(Group, _super);
        function Group() {
            _super.apply(this, arguments);
        }
        return Group;
    })(DataModelWrapper);
    Kurve.Group = Group;
    var Groups = (function (_super) {
        __extends(Groups, _super);
        function Groups() {
            _super.apply(this, arguments);
        }
        return Groups;
    })(DataModelListWrapper);
    Kurve.Groups = Groups;
    var MailFolderDataModel = (function () {
        function MailFolderDataModel() {
        }
        return MailFolderDataModel;
    })();
    Kurve.MailFolderDataModel = MailFolderDataModel;
    var MailFolder = (function (_super) {
        __extends(MailFolder, _super);
        function MailFolder() {
            _super.apply(this, arguments);
        }
        return MailFolder;
    })(DataModelWrapper);
    Kurve.MailFolder = MailFolder;
    var MailFolders = (function (_super) {
        __extends(MailFolders, _super);
        function MailFolders() {
            _super.apply(this, arguments);
        }
        return MailFolders;
    })(DataModelListWrapper);
    Kurve.MailFolders = MailFolders;
    (function (AttachmentType) {
        AttachmentType[AttachmentType["fileAttachment"] = 0] = "fileAttachment";
        AttachmentType[AttachmentType["itemAttachment"] = 1] = "itemAttachment";
        AttachmentType[AttachmentType["referenceAttachment"] = 2] = "referenceAttachment";
    })(Kurve.AttachmentType || (Kurve.AttachmentType = {}));
    var AttachmentType = Kurve.AttachmentType;
    var AttachmentDataModel = (function () {
        function AttachmentDataModel() {
        }
        return AttachmentDataModel;
    })();
    Kurve.AttachmentDataModel = AttachmentDataModel;
    var Attachment = (function (_super) {
        __extends(Attachment, _super);
        function Attachment() {
            _super.apply(this, arguments);
        }
        Attachment.prototype.getType = function () {
            switch (this._data['@odata.type']) {
                case "#microsoft.graph.fileAttachment":
                    return AttachmentType.fileAttachment;
                case "#microsoft.graph.itemAttachment":
                    return AttachmentType.itemAttachment;
                case "#microsoft.graph.referenceAttachment":
                    return AttachmentType.referenceAttachment;
            }
        };
        return Attachment;
    })(DataModelWrapper);
    Kurve.Attachment = Attachment;
    var Attachments = (function (_super) {
        __extends(Attachments, _super);
        function Attachments() {
            _super.apply(this, arguments);
        }
        return Attachments;
    })(DataModelListWrapper);
    Kurve.Attachments = Attachments;
    var Graph = (function () {
        function Graph(identityInfo) {
            this.req = null;
            this.accessToken = null;
            this.KurveIdentity = null;
            this.defaultResourceID = "https://graph.microsoft.com";
            this.baseUrl = "https://graph.microsoft.com/v1.0/";
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            }
            else {
                this.KurveIdentity = identityInfo.identity;
            }
        }
        //Only adds scopes when linked to a v2 Oauth of kurve identity
        Graph.prototype.scopesForV2 = function (scopes) {
            if (!this.KurveIdentity)
                return null;
            if (this.KurveIdentity.getCurrentOauthVersion() === Kurve.OAuthVersion.v1)
                return null;
            else
                return scopes;
        };
        //Users
        Graph.prototype.meAsync = function (odataQuery) {
            var d = new Kurve.Deferred();
            this.me(function (user, error) { return error ? d.reject(error) : d.resolve(user); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.me = function (callback, odataQuery) {
            var scopes = [Scopes.User.Read];
            var urlString = this.buildMeUrl("", odataQuery);
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.userAsync = function (userId, odataQuery, basicProfileOnly) {
            if (basicProfileOnly === void 0) { basicProfileOnly = true; }
            var d = new Kurve.Deferred();
            this.user(userId, function (user, error) { return error ? d.reject(error) : d.resolve(user); }, odataQuery, basicProfileOnly);
            return d.promise;
        };
        Graph.prototype.user = function (userId, callback, odataQuery, basicProfileOnly) {
            if (basicProfileOnly === void 0) { basicProfileOnly = true; }
            var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
            var urlString = this.buildUsersUrl(userId, odataQuery);
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.usersAsync = function (odataQuery, basicProfileOnly) {
            if (basicProfileOnly === void 0) { basicProfileOnly = true; }
            var d = new Kurve.Deferred();
            this.users(function (users, error) { return error ? d.reject(error) : d.resolve(users); }, odataQuery, basicProfileOnly);
            return d.promise;
        };
        Graph.prototype.users = function (callback, odataQuery, basicProfileOnly) {
            if (basicProfileOnly === void 0) { basicProfileOnly = true; }
            var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
            var urlString = this.buildUsersUrl("", odataQuery);
            this.getUsers(urlString, callback, this.scopesForV2(scopes), basicProfileOnly);
        };
        //Groups
        Graph.prototype.groupAsync = function (groupId, odataQuery) {
            var d = new Kurve.Deferred();
            this.group(groupId, function (group, error) { return error ? d.reject(error) : d.resolve(group); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.group = function (groupId, callback, odataQuery) {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildGroupsUrl(groupId, odataQuery);
            this.getGroup(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.groupsAsync = function (odataQuery) {
            var d = new Kurve.Deferred();
            this.groups(function (groups, error) { return error ? d.reject(error) : d.resolve(groups); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.groups = function (callback, odataQuery) {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildGroupsUrl("", odataQuery);
            this.getGroups(urlString, callback, this.scopesForV2(scopes));
        };
        // Messages For User
        Graph.prototype.messageForUserAsync = function (userPrincipalName, messageId, odataQuery) {
            var d = new Kurve.Deferred();
            this.messageForUser(userPrincipalName, messageId, function (message, error) { return error ? d.reject(error) : d.resolve(message); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.messageForUser = function (userPrincipalName, messageId, callback, odataQuery) {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages/" + messageId, odataQuery);
            this.getMessage(urlString, messageId, function (result, error) { return callback(result, error); }, this.scopesForV2(scopes));
        };
        Graph.prototype.messagesForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.messagesForUser(userPrincipalName, function (messages, error) { return error ? d.reject(error) : d.resolve(messages); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.messagesForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages", odataQuery);
            this.getMessages(urlString, function (result, error) { return callback(result, error); }, this.scopesForV2(scopes));
        };
        // MailFolders For User
        Graph.prototype.mailFoldersForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.mailFoldersForUser(userPrincipalName, function (messages, error) { return error ? d.reject(error) : d.resolve(messages); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.mailFoldersForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/mailFolders", odataQuery);
            this.getMailFolders(urlString, function (result, error) { return callback(result, error); }, this.scopesForV2(scopes));
        };
        // Events For User
        Graph.prototype.eventForUserAsync = function (userPrincipalName, eventId, odataQuery) {
            var d = new Kurve.Deferred();
            this.eventForUser(userPrincipalName, eventId, function (event, error) { return error ? d.reject(error) : d.resolve(event); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.eventForUser = function (userPrincipalName, eventId, callback, odataQuery) {
            var scopes = [Scopes.Calendars.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/events/" + eventId, odataQuery);
            this.getEvent(urlString, eventId, function (result, error) { return callback(result, error); }, this.scopesForV2(scopes));
        };
        Graph.prototype.eventsForUserAsync = function (userPrincipalName, endpoint, odataQuery) {
            var d = new Kurve.Deferred();
            this.eventsForUser(userPrincipalName, endpoint, function (events, error) { return error ? d.reject(error) : d.resolve(events); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.eventsForUser = function (userPrincipalName, endpoint, callback, odataQuery) {
            var scopes = [Scopes.Calendars.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/" + EventsEndpoint[endpoint], odataQuery);
            this.getEvents(urlString, endpoint, function (result, error) { return callback(result, error); }, this.scopesForV2(scopes));
        };
        // Groups/Relationships For User
        Graph.prototype.memberOfForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.memberOfForUser(userPrincipalName, function (groups, error) { return error ? d.reject(error) : d.resolve(groups); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.memberOfForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/memberOf", odataQuery);
            this.getGroups(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.managerForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.managerForUser(userPrincipalName, function (user, error) { return error ? d.reject(error) : d.resolve(user); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.managerForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.Directory.ReadAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/manager", odataQuery);
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.directReportsForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.directReportsForUser(userPrincipalName, function (users, error) { return error ? d.reject(error) : d.resolve(users); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.directReportsForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.Directory.ReadAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/directReports", odataQuery);
            this.getUsers(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.profilePhotoForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.profilePhotoForUser(userPrincipalName, function (profilePhoto, error) { return error ? d.reject(error) : d.resolve(profilePhoto); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.profilePhotoForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.User.ReadBasicAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/photo", odataQuery);
            this.getPhoto(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.profilePhotoValueForUserAsync = function (userPrincipalName, odataQuery) {
            var d = new Kurve.Deferred();
            this.profilePhotoValueForUser(userPrincipalName, function (result, error) { return error ? d.reject(error) : d.resolve(result); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.profilePhotoValueForUser = function (userPrincipalName, callback, odataQuery) {
            var scopes = [Scopes.User.ReadBasicAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/photo/$value", odataQuery);
            this.getPhotoValue(urlString, callback, this.scopesForV2(scopes));
        };
        // Message Attachments
        Graph.prototype.messageAttachmentsForUserAsync = function (userPrincipalName, messageId, odataQuery) {
            var d = new Kurve.Deferred();
            this.messageAttachmentsForUser(userPrincipalName, messageId, function (result, error) { return error ? d.reject(error) : d.resolve(result); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.messageAttachmentsForUser = function (userPrincipalName, messageId, callback, odataQuery) {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages/" + messageId + "/attachments", odataQuery);
            this.getMessageAttachments(urlString, callback, this.scopesForV2(scopes));
        };
        Graph.prototype.messageAttachmentForUserAsync = function (userPrincipalName, messageId, attachmentId, odataQuery) {
            var d = new Kurve.Deferred();
            this.messageAttachmentForUser(userPrincipalName, messageId, attachmentId, function (attachment, error) { return error ? d.reject(error) : d.resolve(attachment); }, odataQuery);
            return d.promise;
        };
        Graph.prototype.messageAttachmentForUser = function (userPrincipalName, messageId, attachmentId, callback, odataQuery) {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages/" + messageId + "/attachments/" + attachmentId, odataQuery);
            this.getMessageAttachment(urlString, callback, this.scopesForV2(scopes));
        };
        //http verbs
        Graph.prototype.getAsync = function (url) {
            var d = new Kurve.Deferred();
            this.get(url, function (response, error) { return error ? d.reject(error) : d.resolve(response); });
            return d.promise;
        };
        Graph.prototype.get = function (url, callback, responseType, scopes) {
            var _this = this;
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = function () {
                if (xhr.readyState === 4)
                    if (xhr.status === 200)
                        callback(responseType ? xhr.response : xhr.responseText, null);
                    else
                        callback(null, _this.generateError(xhr));
            };
            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, function (addTokenError) {
                if (addTokenError) {
                    callback(null, addTokenError);
                }
            }, scopes);
        };
        Graph.prototype.generateError = function (xhr) {
            var response = new Kurve.Error();
            response.status = xhr.status;
            response.statusText = xhr.statusText;
            if (xhr.responseType === '' || xhr.responseType === 'text')
                response.text = xhr.responseText;
            else
                response.other = xhr.response;
            return response;
        };
        //Private methods
        Graph.prototype.getUsers = function (urlString, callback, scopes, basicProfileOnly) {
            var _this = this;
            if (basicProfileOnly === void 0) { basicProfileOnly = true; }
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = usersODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var resultsArray = (usersODATA.value ? usersODATA.value : [usersODATA]);
                var users = new Users(_this, resultsArray.map(function (o) { return new User(_this, o); }));
                var nextLink = usersODATA['@odata.nextLink'];
                if (nextLink) {
                    users.nextLink = function (callback) {
                        var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
                        var d = new Kurve.Deferred();
                        _this.getUsers(nextLink, function (result, error) {
                            if (callback)
                                callback(result, error);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, _this.scopesForV2(scopes), basicProfileOnly);
                        return d.promise;
                    };
                }
                callback(users, null);
            }, null, scopes);
        };
        Graph.prototype.getUser = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var userODATA = JSON.parse(result);
                if (userODATA.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = userODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var user = new User(_this, userODATA);
                callback(user, null);
            }, null, scopes);
        };
        Graph.prototype.addAccessTokenAndSend = function (xhr, callback, scopes) {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            }
            else {
                //Using the integrated Identity object
                if (scopes) {
                    //v2 scope based tokens
                    this.KurveIdentity.getAccessTokenForScopes(scopes, false, (function (token, error) {
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
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, (function (token, error) {
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
        };
        Graph.prototype.getMessage = function (urlString, messageId, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Kurve.Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }
                var message = new Message(_this, ODATA);
                callback(message, null);
            }, null, scopes);
        };
        Graph.prototype.getMessages = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var messagesODATA = JSON.parse(result);
                if (messagesODATA.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = messagesODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var resultsArray = (messagesODATA.value ? messagesODATA.value : [messagesODATA]);
                var messages = new Messages(_this, resultsArray.map(function (o) { return new Message(_this, o); }));
                var nextLink = messagesODATA['@odata.nextLink'];
                if (nextLink) {
                    messages.nextLink = function (callback) {
                        var scopes = [Scopes.Mail.Read];
                        var d = new Kurve.Deferred();
                        _this.getMessages(nextLink, function (messages, error) {
                            if (callback)
                                callback(messages, error);
                            else
                                error ? d.reject(error) : d.resolve(messages);
                        }, _this.scopesForV2(scopes));
                        return d.promise;
                    };
                }
                callback(messages, null);
            }, null, scopes);
        };
        Graph.prototype.getEvent = function (urlString, EventId, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Kurve.Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }
                var event = new Event(_this, ODATA);
                callback(event, null);
            }, null, scopes);
        };
        Graph.prototype.getEvents = function (urlString, endpoint, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var odata = JSON.parse(result);
                if (odata.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = odata.error;
                    callback(null, errorODATA);
                    return;
                }
                var resultsArray = (odata.value ? odata.value : [odata]);
                var events = new Events(_this, endpoint, resultsArray.map(function (o) { return new Event(_this, o); }));
                var nextLink = odata['@odata.nextLink'];
                if (nextLink) {
                    events.nextLink = function (callback) {
                        var scopes = [Scopes.Mail.Read];
                        var d = new Kurve.Deferred();
                        _this.getEvents(nextLink, endpoint, function (result, error) {
                            if (callback)
                                callback(result, error);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, _this.scopesForV2(scopes));
                        return d.promise;
                    };
                }
                callback(events, null);
            }, null, scopes);
        };
        Graph.prototype.getGroups = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var groupsODATA = JSON.parse(result);
                if (groupsODATA.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = groupsODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var resultsArray = (groupsODATA.value ? groupsODATA.value : [groupsODATA]);
                var groups = new Groups(_this, resultsArray.map(function (o) { return new Group(_this, o); }));
                var nextLink = groupsODATA['@odata.nextLink'];
                if (nextLink) {
                    groups.nextLink = function (callback) {
                        var scopes = [Scopes.Group.ReadAll];
                        var d = new Kurve.Deferred();
                        _this.getGroups(nextLink, function (result, error) {
                            if (callback)
                                callback(result, error);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, _this.scopesForV2(scopes));
                        return d.promise;
                    };
                }
                callback(groups, null);
            }, null, scopes);
        };
        Graph.prototype.getGroup = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Kurve.Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }
                var group = new Group(_this, ODATA);
                callback(group, null);
            }, null, scopes);
        };
        Graph.prototype.getPhoto = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = ODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var photo = new ProfilePhoto(_this, ODATA);
                callback(photo, null);
            }, null, scopes);
        };
        Graph.prototype.getPhotoValue = function (urlString, callback, scopes) {
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                callback(result, null);
            }, "blob", scopes);
        };
        Graph.prototype.getMailFolders = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var odata = JSON.parse(result);
                if (odata.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = odata.error;
                    callback(null, errorODATA);
                }
                var resultsArray = (odata.value ? odata.value : [odata]);
                var mailFolders = new MailFolders(_this, resultsArray.map(function (o) { return new MailFolder(_this, o); }));
                var nextLink = odata['@odata.nextLink'];
                if (nextLink) {
                    mailFolders.nextLink = function (callback) {
                        var scopes = [Scopes.User.ReadAll];
                        var d = new Kurve.Deferred();
                        _this.getMailFolders(nextLink, function (result, error) {
                            if (callback)
                                callback(result, error);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, _this.scopesForV2(scopes));
                        return d.promise;
                    };
                }
                callback(mailFolders, null);
            }, null, scopes);
        };
        Graph.prototype.getMessageAttachments = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var attachmentsODATA = JSON.parse(result);
                if (attachmentsODATA.error) {
                    var errorODATA = new Kurve.Error();
                    errorODATA.other = attachmentsODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var resultsArray = (attachmentsODATA.value ? attachmentsODATA.value : [attachmentsODATA]);
                var attachments = new Attachments(_this, resultsArray.map(function (o) { return new Attachment(_this, o); }));
                var nextLink = attachmentsODATA['@odata.nextLink'];
                if (nextLink) {
                    attachments.nextLink = function (callback) {
                        var scopes = [Scopes.Mail.Read];
                        var d = new Kurve.Deferred();
                        _this.getMessageAttachments(nextLink, function (attachments, error) {
                            if (callback)
                                callback(attachments, error);
                            else
                                error ? d.reject(error) : d.resolve(attachments);
                        }, _this.scopesForV2(scopes));
                        return d.promise;
                    };
                }
                callback(attachments, null);
            }, null, scopes);
        };
        Graph.prototype.getMessageAttachment = function (urlString, callback, scopes) {
            var _this = this;
            this.get(urlString, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Kurve.Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }
                var attachment = new Attachment(_this, ODATA);
                callback(attachment, null);
            }, null, scopes);
        };
        Graph.prototype.buildUrl = function (root, path, odataQuery) {
            return this.baseUrl + root + path + (odataQuery ? "?" + odataQuery : "");
        };
        Graph.prototype.buildMeUrl = function (path, odataQuery) {
            if (path === void 0) { path = ""; }
            return this.buildUrl("me/", path, odataQuery);
        };
        Graph.prototype.buildUsersUrl = function (path, odataQuery) {
            if (path === void 0) { path = ""; }
            return this.buildUrl("users/", path, odataQuery);
        };
        Graph.prototype.buildGroupsUrl = function (path, odataQuery) {
            if (path === void 0) { path = ""; }
            return this.buildUrl("groups/", path, odataQuery);
        };
        return Graph;
    })();
    Kurve.Graph = Graph;
    var GraphInfo = (function () {
        function GraphInfo(path, scopes, odataQuery) {
            this.path = path;
            this.scopes = scopes;
            this.odataQuery = odataQuery;
        }
        GraphInfo.prototype.getAsync = function (graph) {
            var d = new Kurve.Deferred();
            this.get(graph, function (result, error) { return error ? d.reject(error) : d.resolve(result); });
            return d.promise;
        };
        GraphInfo.prototype.get = function (graph, callback) {
            var url = "https://graph.microsoft.com/v1.0/" + this.path + (this.odataQuery ? "?" + this.odataQuery : "");
            graph.get(url, function (result, errorGet) {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Kurve.Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }
                callback(ODATA, null);
            }, null, graph.scopesForV2(this.scopes));
        };
        return GraphInfo;
    })();
    Kurve.GraphInfo = GraphInfo;
    var ItemAttachmentDataModel = (function () {
        function ItemAttachmentDataModel() {
        }
        return ItemAttachmentDataModel;
    })();
    Kurve.ItemAttachmentDataModel = ItemAttachmentDataModel;
    var ItemAttachment = (function (_super) {
        __extends(ItemAttachment, _super);
        function ItemAttachment() {
            _super.apply(this, arguments);
        }
        ItemAttachment.fromMessageForMe = function (messageId, attachmentId, odataQuery) {
            return new GraphInfo("/me/messages/" + messageId + "/attachments/" + attachmentId, [Scopes.Mail.Read], odataQuery);
        };
        ItemAttachment.fromMessageForUser = function (userId, messageId, attachmentId, odataQuery) {
            return new GraphInfo("/users/#{userId}/messages/" + messageId + "/attachments/" + attachmentId, [Scopes.Mail.Read], odataQuery);
        };
        return ItemAttachment;
    })(ItemAttachmentDataModel);
    Kurve.ItemAttachment = ItemAttachment;
})(Kurve || (Kurve = {}));
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

if (((typeof window != "undefined" && window.module) || (typeof module != "undefined")) && typeof module.exports != "undefined") {
    module.exports = Kurve;
};
