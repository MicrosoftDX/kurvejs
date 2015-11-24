// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Kurve;
(function (Kurve) {
    var Graph = (function () {
        function Graph(tenantId, identityInfo) {
            this.req = null;
            this.state = null;
            this.nonce = null;
            this.accessToken = null;
            this.tenantId = null;
            this.KurveIdentity = null;
            this.defaultResourceID = "https://graph.microsoft.com";
            this.baseUrl = "https://graph.microsoft.com/beta/";
            this.tenantId = tenantId;
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            }
            else {
                this.KurveIdentity = identityInfo.identity;
            }
        }
        //Messages
        //Users
        Graph.prototype.meAsync = function (odataQuery) {
            var d = new Kurve.Deferred();
            this.me(function (user, error) {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(user);
                }
            }, odataQuery);
            return d.promise;
        };
        Graph.prototype.me = function (callback, odataQuery) {
            var urlString = this.buildMeUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUser(urlString, callback);
        };
        Graph.prototype.usersAsync = function (odataQuery) {
            var d = new Kurve.Deferred();
            this.users(function (users, error) {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(users);
                }
            }, odataQuery);
            return d.promise;
        };
        Graph.prototype.users = function (callback, odataQuery) {
            var urlString = this.buildUsersUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUsers(urlString, callback);
        };
        //http verbs
        Graph.prototype.getAsync = function (url) {
            var d = new Kurve.Deferred();
            this.get(url, function (response) {
                d.resolve(response);
            });
            return d.promise;
        };
        Graph.prototype.get = function (url, callback) {
            var xhr = new XMLHttpRequest();
            xhr.onreadystatechange = (function () {
                if (xhr.readyState === 4) {
                    callback(xhr.responseText);
                }
            });
            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr);
        };
        //Private methods
        Graph.prototype.getUsers = function (urlString, callback) {
            var _this = this;
            this.get(urlString, (function (result) {
                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    callback(null, JSON.stringify(usersODATA.error));
                    return;
                }
                var resultsArray = !usersODATA.value ? [usersODATA] : usersODATA.value;
                for (var i = 0; i < resultsArray.length; i++) {
                    _this.decorateUserObject(resultsArray[i]);
                }
                var users = {
                    resultsPage: resultsArray
                };
                //implement nextLink
                var nextLink = usersODATA['@odata.nextLink'];
                if (nextLink) {
                    users.nextLink = (function (callback) {
                        var d = new Kurve.Deferred();
                        _this.getUsers(nextLink, (function (result, error) {
                            if (callback)
                                callback(result, error);
                            else if (error)
                                d.reject(error);
                            else
                                d.resolve(result);
                        }));
                        return d.promise;
                    });
                }
                callback(users, null);
            }));
        };
        Graph.prototype.getUser = function (urlString, callback) {
            var _this = this;
            this.get(urlString, function (result) {
                var userODATA = JSON.parse(result);
                if (userODATA.error) {
                    callback(null, JSON.stringify(userODATA.error));
                    return;
                }
                _this.decorateUserObject(userODATA);
                callback(userODATA, null);
            });
        };
        Graph.prototype.addAccessTokenAndSend = function (xhr) {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            }
            else {
                //Using the integrated Identity object
                this.KurveIdentity.getAccessToken(this.defaultResourceID, (function (token, error) {
                    //cache the token
                    if (error)
                        throw error;
                    xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                    xhr.send();
                }));
            }
        };
        Graph.prototype.decorateUserObject = function (user) {
            var _this = this;
            user.messages = (function (callback, odataQuery) {
                var urlString = _this.buildUsersUrl() + "/" + user.userPrincipalName + "/messages";
                if (odataQuery)
                    urlString += "?" + odataQuery;
                _this.getMessages(urlString, callback, odataQuery);
            });
        };
        Graph.prototype.decorateMessageObject = function (message) {
        };
        Graph.prototype.getMessages = function (urlString, callback, odataQuery) {
            var _this = this;
            var url = urlString;
            if (odataQuery)
                urlString += "?" + odataQuery;
            this.get(url, (function (result) {
                var messagesODATA = JSON.parse(result);
                if (messagesODATA.error) {
                    callback(null, JSON.stringify(messagesODATA.error));
                    return;
                }
                var resultsArray = !messagesODATA.value ? [messagesODATA] : messagesODATA.value;
                for (var i = 0; i < resultsArray.length; i++) {
                    _this.decorateMessageObject(resultsArray[i]);
                }
                var messages = {
                    resultsPage: resultsArray
                };
                var nextLink = messagesODATA['@odata.nextLink'];
                //implement nextLink
                if (nextLink) {
                    messages.nextLink = (function (callback) {
                        var d = new Kurve.Deferred();
                        _this.getMessages(nextLink, (function (result, error) {
                            if (callback)
                                callback(result, error);
                            else if (error)
                                d.reject(error);
                            else
                                d.resolve(result);
                        }));
                        return d.promise;
                    });
                }
                callback(messages, null);
            }));
        };
        Graph.prototype.buildMeUrl = function () {
            return this.baseUrl + "me";
        };
        Graph.prototype.buildUsersUrl = function () {
            return this.baseUrl + this.tenantId + "/users";
        };
        return Graph;
    })();
    Kurve.Graph = Graph;
})(Kurve || (Kurve = {}));
//*********************************************************   
//   
//Kurve js, https://github.com/matvelloso/kurve-js
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
