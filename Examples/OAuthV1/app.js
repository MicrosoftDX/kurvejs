// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Sample;
(function (Sample) {
    var AppV1 = (function () {
        function AppV1() {
            var _this = this;
            //Setup
            this.clientId = document.getElementById("AppID").value;
            this.redirectUri = document.getElementById("redirectUrl").value;
            //Create identity object
            this.identity = new kurve.Identity({
                clientId: this.clientId,
                tokenProcessingUri: this.redirectUri,
                version: kurve.EndPointVersion.v1
            });
            //or  this.identity.loginAsync().then(() => {
            this.identity.loginAsync().then(function () {
                ////Option 1: Manualy passing the access token
                //// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                //this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                //    this.graph = new Kurve.Graph({ defaultAccessToken: token });
                //}));
                //Option 2: Automatically linking to the Identity object
                _this.graph = new kurve.Graph({ identity: _this.identity });
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
        AppV1.prototype.logout = function () {
            this.identity.logOut();
        };
        //Scenario 2: Load users with paging
        AppV1.prototype.loadUsersWithPaging = function () {
            document.getElementById("results").innerHTML = "";
            this.showUsers(this.graph.users, "$top=5");
        };
        //Scenario 3: Load users with custom odata query
        AppV1.prototype.loadUsersWithOdataQuery = function (query) {
            document.getElementById("results").innerHTML = "";
            this.showUsers(this.graph.users, query);
        };
        //Scenario 4: Load user "me"
        AppV1.prototype.loadUserMe = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                return document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            });
        };
        //Scenario 5: Load user by ID
        AppV1.prototype.userById = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.users.$(document.getElementById("userId").value).GetUser()
                .then(function (user) {
                return document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            })
                .fail(function (error) {
                return window.alert(error.status);
            });
        };
        //Scenario 6: Load user "me" and then its messages
        AppV1.prototype.loadUserMessages = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                _this.showMessages(user.self.messages, "$top=2");
            });
        };
        //Scenario 6: Load user "me" and then its messages
        AppV1.prototype.loadUserEvents = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Events:" + "</br>";
                _this.showEvents(user.self.events, "$top=2");
            });
        };
        //Scenario 7: Load user "me" and then its groups
        AppV1.prototype.loadUserGroups = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                _this.showGroups(user.self.memberOf, "$top=5");
            });
        };
        //Scenario 8: Load user "me" and then its manager
        AppV1.prototype.loadUserManager = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                user.self.manager(function (manager) {
                    document.getElementById("results").innerHTML += manager.item.displayName + "</br>";
                });
            });
        };
        //Scenario 9: Load groups with paging
        AppV1.prototype.loadGroupsWithPaging = function () {
            document.getElementById("results").innerHTML = "";
            this.showGroups(this.graph.groups, "$top=5");
        };
        //Scenario 10: Load group by ID
        AppV1.prototype.groupById = function () {
            //document.getElementById("results").innerHTML = "";
            //this.graph.groupAsync((<HTMLInputElement>document.getElementById("groupId")).value).then((group) => {
            //    document.getElementById("results").innerHTML += group.data.displayName + "</br>";
            //});
        };
        //Scenario 11: Load user "me" and then its messages
        AppV1.prototype.loadUserPhoto = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";
                user.self.photo.GetPhotoProperties().then(function (photo) {
                    user.self.GetPhotoImage().then(function (image) {
                        var img = document.createElement("img");
                        var reader = new FileReader();
                        reader.onloadend = function () {
                            img.height = image.item.height;
                            img.width = image.item.width;
                            img.src = reader.result;
                        };
                        reader.readAsDataURL(image.item);
                        document.getElementById("results").appendChild(img);
                    })
                        .fail(function (error) { return window.alert(error.statusText); });
                })
                    .fail(function (error) { return window.alert(error.statusText); });
            })
                .fail(function (error) { return window.alert(error.statusText); });
        };
        //Scenario 12: Is logged in?
        AppV1.prototype.isLoggedIn = function () {
            document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
        };
        //Scenario 13: Who am I?
        AppV1.prototype.whoAmI = function () {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        };
        //--------------------------------Callbacks---------------------------------------------
        AppV1.prototype.showUsers = function (users, odataQuery) {
            var _this = this;
            users.GetUsers(odataQuery)
                .then(function (collection) {
                collection.items.forEach(function (user) {
                    return document.getElementById("results").innerHTML += user.item.displayName + "</br>";
                });
                if (collection.next)
                    _this.showUsers(collection.next);
            })
                .fail(function (error) {
                return document.getElementById("results").innerText = error.statusText;
            });
        };
        AppV1.prototype.showGroups = function (groups, odataQuery) {
            var _this = this;
            groups.GetGroups(odataQuery)
                .then(function (collection) {
                collection.items.forEach(function (group) {
                    return document.getElementById("results").innerHTML += group.item.displayName + "</br>";
                });
                if (collection.next)
                    _this.showUsers(collection.next);
            })
                .fail(function (error) {
                return document.getElementById("results").innerText = error.statusText;
            });
        };
        AppV1.prototype.showMessages = function (messages, odataQuery) {
            var _this = this;
            messages.GetMessages(odataQuery)
                .then(function (collection) {
                collection.items.forEach(function (message) {
                    return document.getElementById("results").innerHTML += message.subject + "</br>";
                });
                if (collection.next)
                    _this.showMessages(collection.next);
            })
                .fail(function (error) {
                return document.getElementById("results").innerText = error.statusText;
            });
        };
        AppV1.prototype.showEvents = function (events, odataQuery) {
            var _this = this;
            events.GetEvents(odataQuery)
                .then(function (collection) {
                collection.items.forEach(function (event) {
                    return document.getElementById("results").innerHTML += event.subject + "</br>";
                });
                if (events.next)
                    _this.showEvents(collection.next);
            })
                .fail(function (error) {
                return document.getElementById("results").innerText = error.statusText;
            });
        };
        return AppV1;
    }());
    Sample.AppV1 = AppV1;
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
