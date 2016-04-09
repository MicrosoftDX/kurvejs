(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else if(typeof exports === 'object')
		exports["Kurve"] = factory();
	else
		root["Kurve"] = factory();
})(this, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;
/******/
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			exports: {},
/******/ 			id: moduleId,
/******/ 			loaded: false
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.loaded = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	function __export(m) {
	    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
	}
	__export(__webpack_require__(1));
	__export(__webpack_require__(3));
	__export(__webpack_require__(4));


/***/ },
/* 1 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var promises_1 = __webpack_require__(2);
	var identity_1 = __webpack_require__(3);
	var requestbuilder_1 = __webpack_require__(4);
	var Graph = (function () {
	    function Graph(identityInfo) {
	        var _this = this;
	        this.req = null;
	        this.accessToken = null;
	        this.KurveIdentity = null;
	        this.defaultResourceID = "https://graph.microsoft.com";
	        this.baseUrl = "https://graph.microsoft.com/v1.0";
	        this.me = function () { return new requestbuilder_1.User(_this, _this.baseUrl); };
	        this.users = function () { return new requestbuilder_1.Users(_this, _this.baseUrl); };
	        if (identityInfo.defaultAccessToken) {
	            this.accessToken = identityInfo.defaultAccessToken;
	        }
	        else {
	            this.KurveIdentity = identityInfo.identity;
	        }
	    }
	    Graph.prototype.Get = function (path, self, scopes) {
	        console.log("GET", path);
	        var d = new promises_1.Deferred();
	        this.get(path, function (error, result) {
	            var jsonResult = JSON.parse(result);
	            if (jsonResult.error) {
	                var errorODATA = new identity_1.Error();
	                errorODATA.other = jsonResult.error;
	                d.reject(errorODATA);
	                return;
	            }
	            d.resolve(new requestbuilder_1.Singleton(jsonResult, self));
	        });
	        return d.promise;
	    };
	    Graph.prototype.GetCollection = function (path, self, next, scopes) {
	        console.log("GET collection", path);
	        var d = new promises_1.Deferred();
	        this.get(path, function (error, result) {
	            var jsonResult = JSON.parse(result);
	            if (jsonResult.error) {
	                var errorODATA = new identity_1.Error();
	                errorODATA.other = jsonResult.error;
	                d.reject(errorODATA);
	                return;
	            }
	            d.resolve(new requestbuilder_1.Collection(jsonResult, self, next));
	        });
	        return d.promise;
	    };
	    Graph.prototype.Post = function (object, path, self, scopes) {
	        console.log("POST", path);
	        var d = new promises_1.Deferred();
	        return d.promise;
	    };
	    Graph.prototype.scopesForV2 = function (scopes) {
	        if (!this.KurveIdentity)
	            return null;
	        if (this.KurveIdentity.getCurrentOauthVersion() === identity_1.OAuthVersion.v1)
	            return null;
	        else
	            return scopes;
	    };
	    Graph.prototype.get = function (url, callback, responseType, scopes) {
	        var _this = this;
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
	    };
	    Graph.prototype.post = function (object, url, callback, responseType, scopes) {
	        var _this = this;
	        var xhr = new XMLHttpRequest();
	        if (responseType)
	            xhr.responseType = responseType;
	        xhr.onreadystatechange = function () {
	            if (xhr.readyState === 4)
	                if (xhr.status === 202)
	                    callback(null, responseType ? xhr.response : xhr.responseText);
	                else
	                    callback(_this.generateError(xhr));
	        };
	        xhr.send(object);
	        xhr.open("GET", url);
	        this.addAccessTokenAndSend(xhr, function (addTokenError) {
	            if (addTokenError) {
	                callback(addTokenError);
	            }
	        }, scopes);
	    };
	    Graph.prototype.generateError = function (xhr) {
	        var response = new identity_1.Error();
	        response.status = xhr.status;
	        response.statusText = xhr.statusText;
	        if (xhr.responseType === '' || xhr.responseType === 'text')
	            response.text = xhr.responseText;
	        else
	            response.other = xhr.response;
	        return response;
	    };
	    Graph.prototype.addAccessTokenAndSend = function (xhr, callback, scopes) {
	        if (this.accessToken) {
	            xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
	            xhr.send();
	        }
	        else {
	            if (scopes) {
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
	                this.KurveIdentity.getAccessToken(this.defaultResourceID, (function (error, token) {
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
	    return Graph;
	}());
	exports.Graph = Graph;


/***/ },
/* 2 */
/***/ function(module, exports) {

	"use strict";
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
	exports.Deferred = Deferred;
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
	exports.Promise = Promise;


/***/ },
/* 3 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var promises_1 = __webpack_require__(2);
	(function (OAuthVersion) {
	    OAuthVersion[OAuthVersion["v1"] = 1] = "v1";
	    OAuthVersion[OAuthVersion["v2"] = 2] = "v2";
	})(exports.OAuthVersion || (exports.OAuthVersion = {}));
	var OAuthVersion = exports.OAuthVersion;
	var Error = (function () {
	    function Error() {
	    }
	    return Error;
	}());
	exports.Error = Error;
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
	exports.IdToken = IdToken;
	var Identity = (function () {
	    function Identity(identitySettings) {
	        var _this = this;
	        this.policy = "";
	        this.clientId = identitySettings.clientId;
	        this.tokenProcessorUrl = identitySettings.tokenProcessingUri;
	        if (identitySettings.version)
	            this.version = identitySettings.version;
	        else
	            this.version = OAuthVersion.v1;
	        this.tokenCache = new TokenCache(identitySettings.tokenStorage);
	        window.addEventListener("message", function (event) {
	            if (event.data.type === "id_token") {
	                if (event.data.error) {
	                    var e = new Error();
	                    e.text = event.data.error;
	                    _this.loginCallback(e);
	                }
	                else {
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
	            if (true) {
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
	        var d = new promises_1.Deferred();
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
	        if (this.version !== OAuthVersion.v1) {
	            var e = new Error();
	            e.statusText = "Currently this identity class is using v2 OAuth mode. You need to use getAccessTokenForScopes() method";
	            callback(e);
	            return;
	        }
	        var token = this.tokenCache.getForResource(resource);
	        if (token) {
	            return callback(null, token.token);
	        }
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
	    };
	    Identity.prototype.getAccessTokenForScopesAsync = function (scopes, promptForConsent) {
	        if (promptForConsent === void 0) { promptForConsent = false; }
	        var d = new promises_1.Deferred();
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
	        this.getTokenCallback = (function (token, error) {
	            if (error) {
	                if (promptForConsent || !error.text) {
	                    callback(null, error);
	                }
	                else if (error.text.indexOf("AADSTS65001") >= 0) {
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
	        var d = new promises_1.Deferred();
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
	        var d = new promises_1.Deferred();
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
	            var redirectUri = (toUrl) ? toUrl : window.location.href.split("#")[0];
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
	}());
	exports.Identity = Identity;


/***/ },
/* 4 */
/***/ function(module, exports) {

	"use strict";
	var __extends = (this && this.__extends) || function (d, b) {
	    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
	    function __() { this.constructor = d; }
	    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	};
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
	exports.OData = OData;
	var pathWithQuery = function (path, odataQuery) {
	    var query = odataQuery && odataQuery.toString();
	    return path + (query ? "?" + query : "");
	};
	var Singleton = (function () {
	    function Singleton(raw, self) {
	        this.raw = raw;
	        this.self = self;
	    }
	    Object.defineProperty(Singleton.prototype, "object", {
	        get: function () {
	            return this.raw;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    return Singleton;
	}());
	exports.Singleton = Singleton;
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
	    Object.defineProperty(Collection.prototype, "objects", {
	        get: function () {
	            return (this.raw.value ? this.raw.value : [this.raw]);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    return Collection;
	}());
	exports.Collection = Collection;
	var Node = (function () {
	    function Node(graph, path) {
	        var _this = this;
	        this.graph = graph;
	        this.path = path;
	        this.pathWithQuery = function (odataQuery, pathSuffix) {
	            if (pathSuffix === void 0) { pathSuffix = ""; }
	            return pathWithQuery(_this.path + pathSuffix, odataQuery);
	        };
	    }
	    return Node;
	}());
	exports.Node = Node;
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
	exports.CollectionNode = CollectionNode;
	var Attachment = (function (_super) {
	    __extends(Attachment, _super);
	    function Attachment(graph, path, attachmentId) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + (attachmentId ? "/" + attachmentId : ""));
	        this.GetAttachment = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this, null); };
	    }
	    return Attachment;
	}(Node));
	exports.Attachment = Attachment;
	var Attachments = (function (_super) {
	    __extends(Attachments, _super);
	    function Attachments(graph, path) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + "/attachments");
	        this.attachment = function (attachmentId) { return new Attachment(_this.graph, _this.path, attachmentId); };
	        this.GetAttachments = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Attachments(_this.graph)); };
	    }
	    return Attachments;
	}(CollectionNode));
	exports.Attachments = Attachments;
	var Message = (function (_super) {
	    __extends(Message, _super);
	    function Message(graph, path, messageId) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + (messageId ? "/" + messageId : ""));
	        this.attachments = function () { return new Attachments(_this.graph, _this.path); };
	        this.GetMessage = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this); };
	        this.SendMessage = function (odataQuery) { return _this.graph.Post(null, _this.pathWithQuery(odataQuery, "/microsoft.graph.sendMail"), _this); };
	    }
	    return Message;
	}(Node));
	exports.Message = Message;
	var Messages = (function (_super) {
	    __extends(Messages, _super);
	    function Messages(graph, path) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + "/messages");
	        this.message = function (messageId) { return new Message(_this.graph, _this.path, messageId); };
	        this.GetMessages = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Messages(_this.graph)); };
	        this.CreateMessage = function (object, odataQuery) { return _this.graph.Post(object, _this.pathWithQuery(odataQuery), _this); };
	    }
	    return Messages;
	}(CollectionNode));
	exports.Messages = Messages;
	var Event = (function (_super) {
	    __extends(Event, _super);
	    function Event(graph, path, eventId) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + (eventId ? "/" + eventId : ""));
	        this.attachments = function () { return new Attachments(_this.graph, _this.path); };
	        this.GetEvent = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this); };
	    }
	    return Event;
	}(Node));
	exports.Event = Event;
	var Events = (function (_super) {
	    __extends(Events, _super);
	    function Events(graph, path) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + "/events");
	        this.event = function (eventId) { return new Event(_this.graph, _this.path, eventId); };
	        this.GetEvents = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Events(_this.graph)); };
	    }
	    return Events;
	}(CollectionNode));
	exports.Events = Events;
	var CalendarView = (function (_super) {
	    __extends(CalendarView, _super);
	    function CalendarView(graph, path) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + "/calendarView");
	        this.GetCalendarView = function (startDate, endDate, odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(queryUnion("startDateTime=" + startDate.toISOString() + "&endDateTime=" + endDate.toISOString(), odataQuery && odataQuery.toString())), _this, new CalendarView(_this.graph)); };
	    }
	    return CalendarView;
	}(CollectionNode));
	exports.CalendarView = CalendarView;
	var MailFolder = (function (_super) {
	    __extends(MailFolder, _super);
	    function MailFolder(graph, path, mailFolderId) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + (mailFolderId ? "/" + mailFolderId : ""));
	        this.GetMailFolder = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this); };
	    }
	    return MailFolder;
	}(Node));
	exports.MailFolder = MailFolder;
	var MailFolders = (function (_super) {
	    __extends(MailFolders, _super);
	    function MailFolders(graph, path) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + "/mailFolders");
	        this.mailFolder = function (mailFolderId) { return new MailFolder(_this.graph, _this.path, mailFolderId); };
	        this.GetMailFolders = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new MailFolders(_this.graph)); };
	    }
	    return MailFolders;
	}(CollectionNode));
	exports.MailFolders = MailFolders;
	var User = (function (_super) {
	    __extends(User, _super);
	    function User(graph, path, userId) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, userId ? path + "/" + userId : path + "/me");
	        this.graph = graph;
	        this.messages = function () { return new Messages(_this.graph, _this.path); };
	        this.events = function () { return new Events(_this.graph, _this.path); };
	        this.calendarView = function () { return new CalendarView(_this.graph, _this.path); };
	        this.mailFolders = function () { return new MailFolders(_this.graph, _this.path); };
	        this.GetUser = function (odataQuery) { return _this.graph.Get(_this.pathWithQuery(odataQuery), _this); };
	    }
	    return User;
	}(Node));
	exports.User = User;
	var Users = (function (_super) {
	    __extends(Users, _super);
	    function Users(graph, path) {
	        var _this = this;
	        if (path === void 0) { path = ""; }
	        _super.call(this, graph, path + "/users");
	        this.user = function (userId) { return new User(_this.graph, _this.path, userId); };
	        this.GetUsers = function (odataQuery) { return _this.graph.GetCollection(_this.pathWithQuery(odataQuery), _this, new Users(_this.graph)); };
	    }
	    return Users;
	}(CollectionNode));
	exports.Users = Users;


/***/ }
/******/ ])
});
;
//# sourceMappingURL=kurve.js.map