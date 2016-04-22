var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var kurve;
(function (kurve) {
    // Adapted from the original source: https://github.com/DirtyHairy/typescript-deferred
    var Error = (function () {
        function Error() {
        }
        return Error;
    }());
    kurve.Error = Error;
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
    }());
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
    }());
    kurve.Deferred = Deferred;
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
    }());
    kurve.Promise = Promise;
    (function (EndPointVersion) {
        EndPointVersion[EndPointVersion["v1"] = 1] = "v1";
        EndPointVersion[EndPointVersion["v2"] = 2] = "v2";
    })(kurve.EndPointVersion || (kurve.EndPointVersion = {}));
    var EndPointVersion = kurve.EndPointVersion;
    (function (Mode) {
        Mode[Mode["Client"] = 1] = "Client";
        Mode[Mode["Node"] = 2] = "Node";
    })(kurve.Mode || (kurve.Mode = {}));
    var Mode = kurve.Mode;
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
    }());
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
    }());
    var IdToken = (function () {
        function IdToken() {
        }
        return IdToken;
    }());
    kurve.IdToken = IdToken;
    var Identity = (function () {
        function Identity(identitySettings) {
            var _this = this;
            this.policy = "";
            this.mode = Mode.Client;
            this.clientId = identitySettings.clientId;
            this.tokenProcessorUrl = identitySettings.tokenProcessingUri;
            //          this.req = new XMLHttpRequest();
            if (identitySettings.version)
                this.version = identitySettings.version;
            else
                this.version = EndPointVersion.v1;
            this.mode = identitySettings.mode;
            if (this.mode === Mode.Client) {
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
        }
        Identity.prototype.parseQueryString = function (str) {
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
        };
        Identity.prototype.token = function (s, url) {
            var start = url.indexOf(s);
            if (start < 0)
                return null;
            var end = url.indexOf("&", start + s.length);
            return url.substring(start, ((end > 0) ? end : url.length));
        };
        Identity.prototype.checkForIdentityRedirect = function () {
            var params = this.parseQueryString(window.location.href);
            var idToken = this.token("#id_token=", window.location.href);
            var accessToken = this.token("#access_token", window.location.href);
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
        Identity.prototype.getCurrentEndPointVersion = function () {
            return this.version;
        };
        Identity.prototype.getAccessTokenAsync = function (resource) {
            var d = new Deferred();
            this.getAccessToken(resource, (function (error, token) {
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
            if (this.version !== EndPointVersion.v1) {
                var e = new Error();
                e.statusText = "Currently this identity class is using v2 OAuth mode. You need to use getAccessTokenForScopes() method";
                callback(e);
                return;
            }
            if (this.mode === Mode.Client) {
                var token = this.tokenCache.getForResource(resource);
                if (token) {
                    return callback(null, token.token);
                }
                //If we got this far, we need to go get this token
                //Need to create the iFrame to invoke the acquire token
                this.getTokenCallback = (function (token, error) {
                    if (error) {
                        callback(error);
                    }
                    else {
                        _this.decodeAccessToken(token, resource);
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
            else {
                var cookies = this.parseNodeCookies(this.req);
                var upn = this.NodeRetrieveDataCallBack("session|" + cookies["kurveSession"]);
                var code = this.NodeRetrieveDataCallBack("code|" + upn);
                var post_data = "grant_type=authorization_code" +
                    "&client_id=" + encodeURIComponent(this.clientId) +
                    "&code=" + encodeURIComponent(code) +
                    "&redirect_uri=" + encodeURIComponent(this.tokenProcessorUrl) +
                    "&resource=" + encodeURIComponent(resource) +
                    "&client_secret=" + encodeURIComponent(this.appSecret);
                var post_options = {
                    host: 'login.microsoftonline.com',
                    port: '443',
                    path: '/common/oauth2/token',
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Content-Length': post_data.length,
                        accept: '*/*'
                    }
                };
                var post_req = this.https.request(post_options, function (response) {
                    response.setEncoding('utf8');
                    response.on('data', function (chunk) {
                        var chunkJson = JSON.parse(chunk);
                        var t = _this.decodeAccessToken(chunkJson.access_token, resource);
                        // this.tokenCache.add(t); //TODO: Persist/retrieve token cache no server
                        callback(null, chunkJson.access_token);
                    });
                });
                post_req.write(post_data);
                post_req.end();
            }
        };
        Identity.prototype.parseNodeCookies = function (req) {
            var list = {};
            var rc = req.headers.cookie;
            rc && rc.split(';').forEach(function (cookie) {
                var parts = cookie.split('=');
                list[parts.shift().trim()] = decodeURI(parts.join('='));
            });
            return list;
        };
        Identity.prototype.handleNodeCallback = function (req, res, https, crypto, persistDataCallback, retrieveDataCallback) {
            var _this = this;
            this.NodePersistDataCallBack = persistDataCallback;
            this.NodeRetrieveDataCallBack = retrieveDataCallback;
            var url = req.url;
            this.req = req;
            this.res = res;
            this.https = https;
            var params = this.parseQueryString(url);
            var code = this.token("code=", url);
            var accessToken = this.token("#access_token", url);
            var cookies = this.parseNodeCookies(req);
            var d = new Deferred();
            if (this.version === EndPointVersion.v1) {
                if (code) {
                    var codeFromRequest = params["code"][0];
                    var stateFromRequest = params["state"][0];
                    var cachedState = retrieveDataCallback("state|" + stateFromRequest);
                    if (cachedState) {
                        if (cachedState === "waiting") {
                            var expiry = new Date(new Date().getTime() + 86400000);
                            persistDataCallback("state|" + stateFromRequest, "done", expiry);
                            var post_data = "grant_type=authorization_code" +
                                "&client_id=" + encodeURIComponent(this.clientId) +
                                "&code=" + encodeURIComponent(codeFromRequest) +
                                "&redirect_uri=" + encodeURIComponent(this.tokenProcessorUrl) +
                                "&resource=" + encodeURIComponent("https://graph.microsoft.com") +
                                "&client_secret=" + encodeURIComponent(this.appSecret);
                            var post_options = {
                                host: 'login.microsoftonline.com',
                                port: '443',
                                path: '/common/oauth2/token',
                                method: 'POST',
                                headers: {
                                    'Content-Type': 'application/x-www-form-urlencoded',
                                    'Content-Length': post_data.length,
                                    accept: '*/*'
                                }
                            };
                            var post_req = https.request(post_options, function (response) {
                                response.setEncoding('utf8');
                                response.on('data', function (chunk) {
                                    var chunkJson = JSON.parse(chunk);
                                    var decodedToken = JSON.parse(_this.base64Decode(chunkJson.access_token.substring(chunkJson.access_token.indexOf('.') + 1, chunkJson.access_token.lastIndexOf('.'))));
                                    var upn = decodedToken.upn;
                                    var sha = crypto.createHash('sha256');
                                    sha.update(Math.random().toString());
                                    var sessionID = sha.digest('hex');
                                    var expiry = new Date(new Date().getTime() + 30 * 60 * 1000);
                                    persistDataCallback("session|" + sessionID, upn, expiry);
                                    persistDataCallback("code|" + upn, codeFromRequest, expiry);
                                    res.writeHead(302, {
                                        'Set-Cookie': 'kurveSession=' + sessionID,
                                        'Location': '/'
                                    });
                                    res.end();
                                    d.resolve(false);
                                });
                            });
                            post_req.write(post_data);
                            post_req.end();
                        }
                        else {
                            //same state has been reused, not allowed
                            res.writeHead(500, "Replay detected", { 'content-type': 'text/plain' });
                            res.end("Replay detected");
                            d.resolve(false);
                        }
                    }
                    else {
                        //state doesn't match any of our cached ones
                        res.writeHead(500, "State doesn't match", { 'content-type': 'text/plain' });
                        res.end("State doesn't match");
                        d.resolve(false);
                    }
                    return d.promise;
                }
                else {
                    if (cookies["kurveSession"]) {
                        var upn = retrieveDataCallback("session|" + cookies["kurveSession"]);
                        if (upn) {
                            d.resolve(true);
                            return d.promise;
                        }
                    }
                    var state = this.generateNonce();
                    var expiry = new Date(new Date().getTime() + 900000);
                    persistDataCallback("state|" + state, "waiting", expiry);
                    var url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&client_id=" +
                        encodeURIComponent(this.clientId) +
                        "&redirect_uri=" + encodeURIComponent(this.tokenProcessorUrl) +
                        "&state=" + encodeURIComponent(state);
                    res.writeHead(302, { 'Location': url });
                    res.end();
                    d.resolve(false);
                    return d.promise;
                }
            }
            else {
                //TODO: v2
                d.resolve(false);
                return d.promise;
            }
        };
        Identity.prototype.getAccessTokenForScopesAsync = function (scopes, promptForConsent) {
            if (promptForConsent === void 0) { promptForConsent = false; }
            var d = new Deferred();
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
            //TODO: Not node compatible
            var d = new Deferred();
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
            //TODO: Not node compatible
            this.loginCallback = callback;
            if (!loginSettings)
                loginSettings = {};
            if (loginSettings.policy)
                this.policy = loginSettings.policy;
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
            //TODO: Not node compatible
            var d = new Deferred();
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
            //TODO: Not node compatible
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
            //TODO: Not node compatible
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
    }());
    kurve.Identity = Identity;
    var Graph = (function () {
        function Graph(identityInfo, mode, https) {
            this.req = null;
            this.accessToken = null;
            this.KurveIdentity = null;
            this.defaultResourceID = "https://graph.microsoft.com";
            this.baseUrl = "https://graph.microsoft.com/v1.0";
            if (https)
                this.https = https;
            this.mode = mode;
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            }
            else {
                this.KurveIdentity = identityInfo.identity;
            }
        }
        Object.defineProperty(Graph.prototype, "me", {
            get: function () { return new User(this, this.baseUrl); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Graph.prototype, "users", {
            get: function () { return new Users(this, this.baseUrl); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Graph.prototype, "groups", {
            get: function () { return new Groups(this, this.baseUrl); },
            enumerable: true,
            configurable: true
        });
        Graph.prototype.Get = function (path, self, scopes, responseType) {
            console.log("GET", path, scopes);
            var d = new Deferred();
            this.get(path, function (error, result) {
                if (!responseType) {
                    var jsonResult = JSON.parse(result);
                    if (jsonResult.error) {
                        var errorODATA = new Error();
                        errorODATA.other = jsonResult.error;
                        d.reject(errorODATA);
                        return;
                    }
                    d.resolve(new Singleton(jsonResult, self));
                }
                else {
                    d.resolve(new Singleton(result, self));
                }
            }, responseType, scopes);
            return d.promise;
        };
        Graph.prototype.GetCollection = function (path, self, next, scopes) {
            console.log("GET collection", path, scopes);
            var d = new Deferred();
            this.get(path, function (error, result) {
                var jsonResult = JSON.parse(result);
                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }
                d.resolve(new Collection(jsonResult, self, next));
            }, null, scopes);
            return d.promise;
        };
        Graph.prototype.Post = function (object, path, self, scopes) {
            console.log("POST", path, scopes);
            var d = new Deferred();
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
        };
        Graph.prototype.get = function (url, callback, responseType, scopes) {
            var _this = this;
            if (this.mode === Mode.Client) {
                var xhr = new XMLHttpRequest();
                if (responseType)
                    xhr.responseType = responseType;
                xhr.onreadystatechange = function () {
                    if (xhr.readyState === 4)
                        if (xhr.status === 200)
                            callback(null, responseType ? xhr.response : xhr.responseText);
                        else
                            callback(_this.generateError(xhr));
                };
                xhr.open("GET", url);
                this.addAccessTokenAndSend(xhr, function (addTokenError) {
                    if (addTokenError) {
                        callback(addTokenError);
                    }
                }, scopes);
            }
            else {
                var token = this.findAccessToken(function (token, error) {
                    var path = url.substr(27, url.length);
                    var options = {
                        host: 'graph.microsoft.com',
                        port: '443',
                        path: path,
                        method: 'GET',
                        headers: {
                            'Content-Type': responseType,
                            accept: '*/*',
                            'Authorization': 'Bearer ' + token
                        }
                    };
                    var post_req = _this.https.request(options, function (response) {
                        response.setEncoding('utf8');
                        response.on('data', function (chunk) {
                            callback(null, chunk);
                        });
                    });
                    post_req.end();
                }, scopes);
            }
        };
        Graph.prototype.findAccessToken = function (callback, scopes) {
            if (this.accessToken) {
                callback(this.accessToken, null);
            }
            else {
                //Using the integrated Identity object
                if (scopes) {
                    //v2 scope based tokens
                    this.KurveIdentity.getAccessTokenForScopes(scopes, false, (function (token, error) {
                        if (error)
                            callback(null, error);
                        else {
                            callback(token, null);
                        }
                    }));
                }
                else {
                    //v1 resource based tokens
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, (function (error, token) {
                        if (error)
                            callback(null, error);
                        else {
                            callback(token, null);
                        }
                    }));
                }
            }
        };
        Graph.prototype.post = function (object, url, callback, responseType, scopes) {
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
        };
        Graph.prototype.generateError = function (xhr) {
            var response = new Error();
            response.status = xhr.status;
            response.statusText = xhr.statusText;
            if (xhr.responseType === '' || xhr.responseType === 'text')
                response.text = xhr.responseText;
            else
                response.other = xhr.response;
            return response;
        };
        Graph.prototype.addAccessTokenAndSend = function (xhr, callback, scopes) {
            this.findAccessToken(function (token, error) {
                if (token) {
                    xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                    xhr.send();
                }
                else {
                    callback(error);
                }
            }, scopes);
        };
        return Graph;
    }());
    kurve.Graph = Graph;
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
    
        ListMessageSubjects(messages:Messages, odata?:string) {
            messages.GetMessages(odata).then(collection => {
                collection.items.forEach(message => console.log(message.subject));
                if (collection.next)
                    ListMessageSubjects(collection.next); // don't need odata after the first time
            })
        }
        ListMessageSubjects(graph.me.messages, new OData().select("subject").toString());
        
    (With async/await support, an iteration pattern can be used intead of recursion)
    
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
    var Scopes = (function () {
        function Scopes() {
        }
        Scopes.rootUrl = "https://graph.microsoft.com/";
        Scopes.General = {
            OpenId: "openid",
            OfflineAccess: "offline_access",
        };
        Scopes.User = {
            Read: Scopes.rootUrl + "User.Read",
            ReadAll: Scopes.rootUrl + "User.Read.All",
            ReadWrite: Scopes.rootUrl + "User.ReadWrite",
            ReadWriteAll: Scopes.rootUrl + "User.ReadWrite.All",
            ReadBasicAll: Scopes.rootUrl + "User.ReadBasic.All",
        };
        Scopes.Contacts = {
            Read: Scopes.rootUrl + "Contacts.Read",
            ReadWrite: Scopes.rootUrl + "Contacts.ReadWrite",
        };
        Scopes.Directory = {
            ReadAll: Scopes.rootUrl + "Directory.Read.All",
            ReadWriteAll: Scopes.rootUrl + "Directory.ReadWrite.All",
            AccessAsUserAll: Scopes.rootUrl + "Directory.AccessAsUser.All",
        };
        Scopes.Group = {
            ReadAll: Scopes.rootUrl + "Group.Read.All",
            ReadWriteAll: Scopes.rootUrl + "Group.ReadWrite.All",
            AccessAsUserAll: Scopes.rootUrl + "Directory.AccessAsUser.All"
        };
        Scopes.Mail = {
            Read: Scopes.rootUrl + "Mail.Read",
            ReadWrite: Scopes.rootUrl + "Mail.ReadWrite",
            Send: Scopes.rootUrl + "Mail.Send",
        };
        Scopes.Calendars = {
            Read: Scopes.rootUrl + "Calendars.Read",
            ReadWrite: Scopes.rootUrl + "Calendars.ReadWrite",
        };
        Scopes.Files = {
            Read: Scopes.rootUrl + "Files.Read",
            ReadAll: Scopes.rootUrl + "Files.Read.All",
            ReadWrite: Scopes.rootUrl + "Files.ReadWrite",
            ReadWriteAppFolder: Scopes.rootUrl + "Files.ReadWrite.AppFolder",
            ReadWriteSelected: Scopes.rootUrl + "Files.ReadWrite.Selected",
        };
        Scopes.Tasks = {
            ReadWrite: Scopes.rootUrl + "Tasks.ReadWrite",
        };
        Scopes.People = {
            Read: Scopes.rootUrl + "People.Read",
            ReadWrite: Scopes.rootUrl + "People.ReadWrite",
        };
        Scopes.Notes = {
            Create: Scopes.rootUrl + "Notes.Create",
            ReadWriteCreatedByApp: Scopes.rootUrl + "Notes.ReadWrite.CreatedByApp",
            Read: Scopes.rootUrl + "Notes.Read",
            ReadAll: Scopes.rootUrl + "Notes.Read.All",
            ReadWriteAll: Scopes.rootUrl + "Notes.ReadWrite.All",
        };
        return Scopes;
    }());
    kurve.Scopes = Scopes;
    var queryUnion = function (query1, query2) { return (query1 ? query1 + (query2 ? "&" + query2 : "") : query2); };
    var OData = (function () {
        function OData(query) {
            var _this = this;
            this.query = query;
            this.toString = function () { return _this.query; };
            this.odata = function (query) {
                _this.query = queryUnion(_this.query, query);
                return _this;
            };
            this.select = function () {
                var fields = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    fields[_i - 0] = arguments[_i];
                }
                return _this.odata("$select=" + fields.join(","));
            };
            this.expand = function () {
                var fields = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    fields[_i - 0] = arguments[_i];
                }
                return _this.odata("$expand=" + fields.join(","));
            };
            this.filter = function (query) { return _this.odata("$filter=" + query); };
            this.orderby = function () {
                var fields = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    fields[_i - 0] = arguments[_i];
                }
                return _this.odata("$orderby=" + fields.join(","));
            };
            this.top = function (items) { return _this.odata("$top=" + items); };
            this.skip = function (items) { return _this.odata("$skip=" + items); };
        }
        return OData;
    }());
    kurve.OData = OData;
    var pathWithQuery = function (path, odataQuery) {
        var query = odataQuery && odataQuery.toString();
        return path + (query ? "?" + query : "");
    };
    var Singleton = (function () {
        function Singleton(raw, self) {
            this.raw = raw;
            this.self = self;
        }
        Object.defineProperty(Singleton.prototype, "item", {
            get: function () {
                return this.raw;
            },
            enumerable: true,
            configurable: true
        });
        return Singleton;
    }());
    kurve.Singleton = Singleton;
    var Collection = (function () {
        function Collection(raw, self, next) {
            this.raw = raw;
            this.self = self;
            this.next = next;
            var nextLink = this.raw["@odata.nextLink"];
            if (nextLink) {
                this.next.nextLink = nextLink;
            }
            else {
                this.next = undefined;
            }
        }
        Object.defineProperty(Collection.prototype, "items", {
            get: function () {
                return (this.raw.value ? this.raw.value : [this.raw]);
            },
            enumerable: true,
            configurable: true
        });
        return Collection;
    }());
    kurve.Collection = Collection;
    var Node = (function () {
        function Node(graph, path) {
            var _this = this;
            this.graph = graph;
            this.path = path;
            //Only adds scopes when linked to a v2 Oauth of kurve identity
            this.scopesForV2 = function (scopes) {
                return _this.graph.KurveIdentity && _this.graph.KurveIdentity.getCurrentEndPointVersion() === EndPointVersion.v2 ? scopes : null;
            };
            this.pathWithQuery = function (odataQuery, pathSuffix) {
                if (pathSuffix === void 0) { pathSuffix = ""; }
                return pathWithQuery(_this.path + pathSuffix, odataQuery);
            };
        }
        return Node;
    }());
    kurve.Node = Node;
    var CollectionNode = (function (_super) {
        __extends(CollectionNode, _super);
        function CollectionNode() {
            var _this = this;
            _super.apply(this, arguments);
            this.pathWithQuery = function (odataQuery, pathSuffix) {
                if (pathSuffix === void 0) { pathSuffix = ""; }
                return _this._nextLink || pathWithQuery(_this.path + pathSuffix, odataQuery);
            };
        }
        Object.defineProperty(CollectionNode.prototype, "nextLink", {
            set: function (pathWithQuery) {
                this._nextLink = pathWithQuery;
            },
            enumerable: true,
            configurable: true
        });
        return CollectionNode;
    }(Node));
    kurve.CollectionNode = CollectionNode;
    var Attachment = (function (_super) {
        __extends(Attachment, _super);
        function Attachment(graph, path, context, attachmentId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + (attachmentId ? "/" + attachmentId : ""));
            this.context = context;
            this.GetAttachment = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2(Attachment.scopes[_this.context])); };
        }
        Attachment.scopes = {
            messages: [Scopes.Mail.Read],
            events: [Scopes.Calendars.Read],
        };
        return Attachment;
    }(Node));
    kurve.Attachment = Attachment;
    var Attachments = (function (_super) {
        __extends(Attachments, _super);
        function Attachments(graph, path, context) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/attachments");
            this.context = context;
            this.$ = function (attachmentId) { return new Attachment(_this.graph, _this.path, _this.context, attachmentId); };
            this.GetAttachments = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Attachments(_this.graph, null, _this.context), _this.scopesForV2(Attachment.scopes[_this.context])); };
        }
        return Attachments;
    }(CollectionNode));
    kurve.Attachments = Attachments;
    var Message = (function (_super) {
        __extends(Message, _super);
        function Message(graph, path, messageId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + (messageId ? "/" + messageId : ""));
            this.GetMessage = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.Mail.Read])); };
            this.SendMessage = function (odataQuery) { return _this.graph.Post(null, _this.pathWithQuery(odataQuery, "/microsoft.graph.sendMail"), _this, _this.scopesForV2([Scopes.Mail.Send])); };
        }
        Object.defineProperty(Message.prototype, "attachments", {
            get: function () { return new Attachments(this.graph, this.path, "messages"); },
            enumerable: true,
            configurable: true
        });
        return Message;
    }(Node));
    kurve.Message = Message;
    var Messages = (function (_super) {
        __extends(Messages, _super);
        function Messages(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/messages");
            this.$ = function (messageId) { return new Message(_this.graph, _this.path, messageId); };
            this.GetMessages = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Messages(_this.graph), _this.scopesForV2([Scopes.Mail.Read])); };
            this.CreateMessage = function (object, odataQuery) { return _this.graph.Post(object, _this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.Mail.ReadWrite])); };
        }
        return Messages;
    }(CollectionNode));
    kurve.Messages = Messages;
    var Event = (function (_super) {
        __extends(Event, _super);
        function Event(graph, path, eventId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + (eventId ? "/" + eventId : ""));
            this.GetEvent = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.Calendars.Read])); };
        }
        Object.defineProperty(Event.prototype, "attachments", {
            get: function () { return new Attachments(this.graph, this.path, "events"); },
            enumerable: true,
            configurable: true
        });
        return Event;
    }(Node));
    kurve.Event = Event;
    var Events = (function (_super) {
        __extends(Events, _super);
        function Events(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/events");
            this.$ = function (eventId) { return new Event(_this.graph, _this.path, eventId); };
            this.GetEvents = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Events(_this.graph), _this.scopesForV2([Scopes.Calendars.Read])); };
        }
        return Events;
    }(CollectionNode));
    kurve.Events = Events;
    var CalendarView = (function (_super) {
        __extends(CalendarView, _super);
        function CalendarView(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/calendarView");
            this.dateRange = function (startDate, endDate) { return ("startDateTime=" + startDate.toISOString() + "&endDateTime=" + endDate.toISOString()); };
            this.GetCalendarView = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new CalendarView(_this.graph), _this.scopesForV2([Scopes.Calendars.Read])); };
        }
        return CalendarView;
    }(CollectionNode));
    kurve.CalendarView = CalendarView;
    var MailFolder = (function (_super) {
        __extends(MailFolder, _super);
        function MailFolder(graph, path, mailFolderId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + (mailFolderId ? "/" + mailFolderId : ""));
            this.GetMailFolder = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.Mail.Read])); };
        }
        return MailFolder;
    }(Node));
    kurve.MailFolder = MailFolder;
    var MailFolders = (function (_super) {
        __extends(MailFolders, _super);
        function MailFolders(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/mailFolders");
            this.$ = function (mailFolderId) { return new MailFolder(_this.graph, _this.path, mailFolderId); };
            this.GetMailFolders = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new MailFolders(_this.graph), _this.scopesForV2([Scopes.Mail.Read])); };
        }
        return MailFolders;
    }(CollectionNode));
    kurve.MailFolders = MailFolders;
    var Photo = (function (_super) {
        __extends(Photo, _super);
        function Photo(graph, path, context) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/photo");
            this.context = context;
            this.GetPhotoProperties = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2(Photo.scopes[_this.context])); };
            this.GetPhotoImage = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery, "/$value"), _this, _this.scopesForV2(Photo.scopes[_this.context]), "blob"); };
        }
        Photo.scopes = {
            user: [Scopes.User.ReadBasicAll],
            group: [Scopes.Group.ReadAll],
            contact: [Scopes.Contacts.Read]
        };
        return Photo;
    }(Node));
    kurve.Photo = Photo;
    var Manager = (function (_super) {
        __extends(Manager, _super);
        function Manager(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/manager");
            this.GetManager = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.User.ReadAll])); };
        }
        return Manager;
    }(Node));
    kurve.Manager = Manager;
    var MemberOf = (function (_super) {
        __extends(MemberOf, _super);
        function MemberOf(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/memberOf");
            this.GetGroups = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new MemberOf(_this.graph), _this.scopesForV2([Scopes.User.ReadAll])); };
        }
        return MemberOf;
    }(CollectionNode));
    kurve.MemberOf = MemberOf;
    var DirectReport = (function (_super) {
        __extends(DirectReport, _super);
        function DirectReport(graph, path, userId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/" + userId);
            this.graph = graph;
            this.GetDirectReport = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.User.Read])); };
        }
        return DirectReport;
    }(Node));
    kurve.DirectReport = DirectReport;
    var DirectReports = (function (_super) {
        __extends(DirectReports, _super);
        function DirectReports(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/directReports");
            this.$ = function (userId) { return new DirectReport(_this.graph, _this.path, userId); };
            this.GetDirectReports = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new DirectReports(_this.graph), _this.scopesForV2([Scopes.User.Read])); };
        }
        return DirectReports;
    }(CollectionNode));
    kurve.DirectReports = DirectReports;
    var User = (function (_super) {
        __extends(User, _super);
        function User(graph, path, userId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, userId ? path + "/" + userId : path + "/me");
            this.graph = graph;
            this.GetUser = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.User.Read])); };
        }
        Object.defineProperty(User.prototype, "messages", {
            get: function () { return new Messages(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "events", {
            get: function () { return new Events(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "calendarView", {
            get: function () { return new CalendarView(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "mailFolders", {
            get: function () { return new MailFolders(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "photo", {
            get: function () { return new Photo(this.graph, this.path, "user"); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "manager", {
            get: function () { return new Manager(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "directReports", {
            get: function () { return new DirectReports(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(User.prototype, "memberOf", {
            get: function () { return new MemberOf(this.graph, this.path); },
            enumerable: true,
            configurable: true
        });
        return User;
    }(Node));
    kurve.User = User;
    var Users = (function (_super) {
        __extends(Users, _super);
        function Users(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/users");
            this.$ = function (userId) { return new User(_this.graph, _this.path, userId); };
            this.GetUsers = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Users(_this.graph), _this.scopesForV2([Scopes.User.Read])); };
        }
        return Users;
    }(CollectionNode));
    kurve.Users = Users;
    var Group = (function (_super) {
        __extends(Group, _super);
        function Group(graph, path, groupId) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/" + groupId);
            this.graph = graph;
            this.GetGroup = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, _this.scopesForV2([Scopes.Group.ReadAll])); };
        }
        return Group;
    }(Node));
    kurve.Group = Group;
    var Groups = (function (_super) {
        __extends(Groups, _super);
        function Groups(graph, path) {
            var _this = this;
            if (path === void 0) { path = ""; }
            _super.call(this, graph, path + "/groups");
            this.$ = function (groupId) { return new Group(_this.graph, _this.path, groupId); };
            this.GetGroups = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Groups(_this.graph), _this.scopesForV2([Scopes.Group.ReadAll])); };
        }
        return Groups;
    }(CollectionNode));
    kurve.Groups = Groups;
})(kurve || (kurve = {}));
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
