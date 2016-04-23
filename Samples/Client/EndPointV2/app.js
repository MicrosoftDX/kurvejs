// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Sample;
(function (Sample) {
    var Error = (function () {
        function Error() {
        }
        return Error;
    }());
    var AppV2 = (function () {
        function AppV2() {
            var _this = this;
            //Setup
            this.clientId = document.getElementById("AppID").value;
            this.redirectUri = document.getElementById("redirectUrl").value;
            //Create identity object
            this.identity = new kurve.Identity({
                clientId: this.clientId,
                tokenProcessingUri: this.redirectUri,
                version: kurve.EndPointVersion.v2
            });
            //We can request for specific scopes during logon (so user will have to consent them right away and not during the flow of the app
            //The list of available consents is available under Kuve.Scopes module
            this.identity.loginAsync({ scopes: [kurve.Scopes.Mail.Read, kurve.Scopes.General.OpenId] }).then(function () {
                _this.graph = new kurve.Graph({ identity: _this.identity });
                //Update UI
                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (function () { _this.logout(); }));
                document.getElementById("requestAccessToken").addEventListener("click", (function () { _this.requestAccessToken(); }));
                document.getElementById("meUser").addEventListener("click", (function () { _this.loadUserMe(); }));
                document.getElementById("userMessages").addEventListener("click", (function () { _this.loadUserMessages(); }));
                document.getElementById("userEvents").addEventListener("click", (function () { _this.loadUserEvents(); }));
                document.getElementById("userPhoto").addEventListener("click", (function () { _this.loadUserPhoto(); }));
                document.getElementById("loggedIn").addEventListener("click", (function () { _this.isLoggedIn(); }));
                document.getElementById("whoAmI").addEventListener("click", (function () { _this.whoAmI(); }));
            });
        }
        //-----------------------------------------------Scenarios---------------------------------------------
        //Scenario 1: Logout
        AppV2.prototype.logout = function () {
            this.identity.logOut();
        };
        //Scenario 2: Request access token
        AppV2.prototype.requestAccessToken = function () {
            document.getElementById("results").innerHTML = "";
            //Using false for prompt will cause it to attempt to silently get the token. If it fails, it will then re-attempt, this time
            //Prompting for a UI for consent
            this.identity.getAccessTokenForScopesAsync([kurve.Scopes.User.Read, kurve.Scopes.Mail.Read], false).then(function (token) {
                document.getElementById("results").innerText = "token: " + token;
            });
        };
        //Scenario 3: Load user "me"
        AppV2.prototype.loadUserMe = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            });
        };
        //Scenario 4: Load user "me" and then its messages
        AppV2.prototype.loadUserMessages = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.messages.GetMessages("$top=2").then(function (messages) {
                _this.messagesCallback(messages, null);
            }).fail(function (error) {
                _this.messagesCallback(null, error);
            });
        };
        AppV2.prototype.loadUserEvents = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.self.events.GetEvents("$top=2").then(function (events) {
                    _this.eventsCallback(events, null);
                }).fail(function (error) {
                    _this.eventsCallback(null, error);
                });
            });
        };
        //Scenario 5: Load user "me" and then its messages
        AppV2.prototype.loadUserPhoto = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";
                user.self.photo.GetPhotoProperties().then(function (photo) {
                    var x = photo;
                });
                user.self.photo.GetPhotoImage().then(function (photoValue) {
                    var img = document.createElement("img");
                    var reader = new FileReader();
                    reader.onloadend = function () {
                        img.src = reader.result;
                    };
                    reader.readAsDataURL(photoValue.item);
                    document.getElementById("results").appendChild(img);
                });
            });
        };
        //Scenario 6: Is logged in?
        AppV2.prototype.isLoggedIn = function () {
            document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
        };
        //Scenario 7: Who am I?
        AppV2.prototype.whoAmI = function () {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        };
        //--------------------------------Callbacks---------------------------------------------
        AppV2.prototype.messagesCallback = function (messages, error) {
            var _this = this;
            if (messages.items) {
                messages.items.forEach(function (item) {
                    document.getElementById("results").innerHTML += item.subject + "</br>";
                });
            }
            if (messages.next) {
                messages.next.GetMessages().then(function (messages) {
                    _this.messagesCallback(messages, error);
                });
            }
        };
        AppV2.prototype.eventsCallback = function (events, error) {
            var _this = this;
            if (events.items) {
                events.items.forEach(function (event) {
                    document.getElementById("results").innerHTML += event.subject + "</br>";
                });
            }
            if (events.next) {
                events.next.GetEvents().then(function (stuff) {
                    _this.eventsCallback(stuff, error);
                });
            }
        };
        return AppV2;
    }());
    Sample.AppV2 = AppV2;
})(Sample || (Sample = {}));
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
