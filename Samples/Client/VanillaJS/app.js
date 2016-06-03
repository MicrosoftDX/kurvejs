/// <reference path="../../../dist/kurve.d.ts" />
var kurve = window["Kurve"];
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var init = function () { return new App(); };
var App = (function () {
    function App() {
        var _this = this;
        // Setup
        var model = document.getElementById("model").value == 'v2' ? kurve.EndPointVersion.v2 : kurve.EndPointVersion.v1;
        var clientId = document.getElementById("AppID").value || model == kurve.EndPointVersion.v2 ? "13c5e4af-5ea6-4b48-8989-ca25c96ba1c4" : "636e98ea-3024-4810-a66e-cda4bfa0a489";
        var loc = document.URL;
        var redirectUri = loc.substr(0, loc.indexOf("/Samples/Client/VanillaJS")) + "/dist/login.html";
        // Create identity object
        this.identity = new kurve.Identity({
            clientId: clientId,
            tokenProcessingUri: redirectUri,
            version: model,
            mode: kurve.Mode.Client
        });
        var scopes = model == kurve.EndPointVersion.v2 ? { scopes: [kurve.Scopes.Mail.Read, kurve.Scopes.General.OpenId] } : {};
        // Login
        this.identity.loginAsync(scopes).then(function (_) {
            ////Option 1: Manualy passing the access token
            //// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
            //this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
            //    this.graph = new Kurve.Graph({ defaultAccessToken: token });
            //}));
            //Option 2: Automatically linking to the Identity object
            _this.graph = new kurve.Graph({ identity: _this.identity }, kurve.Mode.Client);
            //Update UI
            document.getElementById("initDiv").style.display = "none";
            document.getElementById("scenarios").style.display = "";
            document.getElementById("logoutBtn").addEventListener("click", (function () { _this.logout(); }));
            document.getElementById("usersWithPaging").addEventListener("click", (function () { _this.loadUsersWithPaging(); }));
            document.getElementById("usersWithCustomODATA").addEventListener("click", (function () { _this.loadUsersWithOdataQuery(document.getElementById('odataquery').value); }));
            document.getElementById("meUser").addEventListener("click", (function () { _this.loadUserMe(); }));
            document.getElementById("userById").addEventListener("click", (function () { _this.userById(); }));
            document.getElementById("userMessages").addEventListener("click", (function () { _this.loadUserMessages(); }));
            document.getElementById("userEvents").addEventListener("click", (function () { _this.loadUserEvents(); }));
            document.getElementById("userGroups").addEventListener("click", (function () { _this.loadUserGroups(); }));
            document.getElementById("userManager").addEventListener("click", (function () { _this.loadUserManager(); }));
            document.getElementById("groupsWithPaging").addEventListener("click", (function () { _this.loadGroupsWithPaging(); }));
            document.getElementById("groupById").addEventListener("click", (function () { _this.groupById(); }));
            document.getElementById("userPhoto").addEventListener("click", (function () { _this.loadUserPhoto(); }));
            document.getElementById("loggedIn").addEventListener("click", (function () { _this.isLoggedIn(); }));
            document.getElementById("whoAmI").addEventListener("click", (function () { _this.whoAmI(); }));
        });
    }
    //-----------------------------------------------Scenarios---------------------------------------------
    //Scenario 1: Logout
    App.prototype.logout = function () {
        this.identity.logOut();
    };
    //Scenario 2: Load users with paging
    App.prototype.loadUsersWithPaging = function () {
        var _this = this;
        document.getElementById("results").innerHTML = "";
        this.graph.users.GetUsers().then(function (users) {
            return _this.showUsers(users);
        });
    };
    //Scenario 3: Load users with custom odata query
    App.prototype.loadUsersWithOdataQuery = function (query) {
        var _this = this;
        document.getElementById("results").innerHTML = "";
        this.graph.users.GetUsers(query).then(function (users) {
            return _this.showUsers(users);
        });
    };
    //Scenario 4: Load user "me"
    App.prototype.loadUserMe = function () {
        document.getElementById("results").innerHTML = "";
        this.graph.me.GetUser().then(function (user) {
            return document.getElementById("results").innerHTML += user.displayName + "</br>";
        }).fail(function (error) {
            return document.getElementById("results").innerText = JSON.stringify(error);
        });
    };
    //Scenario 5: Load user by ID
    App.prototype.userById = function () {
        document.getElementById("results").innerHTML = "";
        this.graph.users.$(document.getElementById("userId").value).GetUser().then(function (user) {
            document.getElementById("results").innerHTML += user.displayName + "</br>";
        }).fail(function (error) {
            window.alert(error.status);
        });
    };
    //Scenario 6: Load user "me" and then its messages
    App.prototype.loadUserMessages = function () {
        var _this = this;
        document.getElementById("results").innerHTML = "";
        this.graph.me.GetUser().then(function (user) {
            document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
            document.getElementById("results").innerHTML += "Messages:" + "</br>";
            user._context.messages.GetMessages().then(function (messages) {
                return _this.showMessages(messages);
            });
        });
    };
    //Scenario 6: Load user "me" and then its events
    App.prototype.loadUserEvents = function () {
        var _this = this;
        document.getElementById("results").innerHTML = "";
        this.graph.me.GetUser().then(function (user) {
            document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
            document.getElementById("results").innerHTML += "Events:" + "</br>";
            user._context.events.GetEvents().then(function (events) {
                return _this.showEvents(events);
            });
        });
    };
    //Scenario 7: Load user "me" and then its groups
    App.prototype.loadUserGroups = function () {
        var _this = this;
        document.getElementById("results").innerHTML = "";
        this.graph.me.GetUser().then(function (user) {
            document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
            document.getElementById("results").innerHTML += "Groups:" + "</br>";
            user._context.memberOf.GetGroups("$top=5").then(function (groups) {
                return _this.showGroups(groups);
            });
        });
    };
    //Scenario 8: Load user "me" and then its manager
    App.prototype.loadUserManager = function () {
        document.getElementById("results").innerHTML = "";
        this.graph.me.GetUser().then(function (user) {
            document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
            document.getElementById("results").innerHTML += "Manager:" + "</br>";
            user._context.manager.GetUser().then(function (manager) {
                document.getElementById("results").innerHTML += manager.displayName + "</br>";
            });
        });
    };
    //Scenario 9: Load groups with paging
    App.prototype.loadGroupsWithPaging = function () {
        var _this = this;
        document.getElementById("results").innerHTML = "";
        this.graph.groups.GetGroups().then(function (groups) {
            return _this.showGroups(groups, 5);
        });
    };
    //Scenario 10: Load group by ID
    App.prototype.groupById = function () {
        document.getElementById("results").innerHTML = "";
        this.graph.groups.$(document.getElementById("groupId").value).GetGroup().then(function (group) {
            document.getElementById("results").innerHTML += group.displayName + "</br>";
        });
    };
    //Scenario 11: Load user "me" and then its messages
    App.prototype.loadUserPhoto = function () {
        document.getElementById("results").innerHTML = "";
        this.graph.me.GetUser().then(function (user) {
            document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
            document.getElementById("results").innerHTML += "Photo:" + "</br>";
            user._context.photo.GetPhotoProperties().then(function (photo) {
                //Photo metadata
                var x = photo;
            });
            user._context.photo.GetPhotoImage().then(function (photoValue) {
                var img = document.createElement("img");
                var reader = new FileReader();
                reader.onloadend = function () {
                    img.src = reader.result;
                };
                reader.readAsDataURL(photoValue.raw);
                document.getElementById("results").appendChild(img);
            });
        });
    };
    //Scenario 12: Is logged in?
    App.prototype.isLoggedIn = function () {
        document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
    };
    //Scenario 13: Who am I?
    App.prototype.whoAmI = function () {
        document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
    };
    //--------------------------------Callbacks---------------------------------------------
    App.prototype.showUsers = function (users, limit) {
        var _this = this;
        if (limit === void 0) { limit = 5; }
        for (var _i = 0, _a = users.value; _i < _a.length; _i++) {
            var user = _a[_i];
            document.getElementById("results").innerHTML += user.displayName + "</br>";
            if (limit-- <= 0)
                return;
        }
        users._next && users._next().then(function (nextUsers) {
            return _this.showUsers(nextUsers, limit);
        }).fail(function (error) {
            return document.getElementById("results").innerText = JSON.stringify(error);
        });
    };
    App.prototype.showGroups = function (groups, limit) {
        var _this = this;
        if (limit === void 0) { limit = 5; }
        for (var _i = 0, _a = groups.value; _i < _a.length; _i++) {
            var group = _a[_i];
            document.getElementById("results").innerHTML += group.displayName + "</br>";
            if (limit-- <= 0)
                return;
        }
        groups._next && groups._next().then(function (nextGroups) {
            return _this.showGroups(nextGroups);
        }).fail(function (error) {
            return document.getElementById("results").innerText = JSON.stringify(error);
        });
    };
    App.prototype.showMessages = function (messages, limit) {
        var _this = this;
        if (limit === void 0) { limit = 5; }
        for (var _i = 0, _a = messages.value; _i < _a.length; _i++) {
            var message = _a[_i];
            document.getElementById("results").innerHTML += message.subject + "</br>";
            if (limit-- <= 0)
                return;
        }
        messages._next && messages._next().then(function (nextMessages) {
            return _this.showMessages(nextMessages, limit);
        }).fail(function (error) {
            return document.getElementById("results").innerText = error.statusText;
        });
    };
    App.prototype.showEvents = function (events, limit) {
        var _this = this;
        if (limit === void 0) { limit = 5; }
        for (var _i = 0, _a = events.value; _i < _a.length; _i++) {
            var event_1 = _a[_i];
            document.getElementById("results").innerHTML += event_1.subject + "</br>";
            if (limit-- <= 0)
                return;
        }
        events._next && events._next().then(function (nextEvents) {
            return _this.showEvents(nextEvents);
        }).fail(function (error) {
            return document.getElementById("results").innerText = error.statusText;
        });
    };
    return App;
}());
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
