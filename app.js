// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Sample;
(function (Sample) {
    var Error = (function () {
        function Error() {
        }
        return Error;
    })();
    var App = (function () {
        //constructor() {
        //    this.test();
        //}
        function App() {
            //Setup
            this.clientId = document.getElementById("classID").value;
            this.redirectUri = document.getElementById("redirectUrl").value;
            //Create identity object
            this.identity = new Kurve.Identity(this.clientId, this.redirectUri);
        }
        App.prototype.doLogin = function () {
            var _this = this;
            this.identity.loginAsync(window.location.href)
                .then(function () {
                //////Option 1: Manualy passing the access token
                ////// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                ////this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                ////    this.graph = new Kurve.Graph({ defaultAccessToken: token });
                ////}));
                //Option 2: Automatically linking to the Identity object
                _this.graph = new Kurve.Graph({ identity: _this.identity });
                //Update UI
                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (function () { _this.logout(); }));
                document.getElementById("usersWithPaging").addEventListener("click", (function () { _this.loadUsersWithPaging(); }));
                document.getElementById("usersWithCustomODATA").addEventListener("click", (function () { _this.loadUsersWithOdataQuery(document.getElementById('odataquery').value); }));
                document.getElementById("meUser").addEventListener("click", (function () { _this.loadUserMe(); }));
                document.getElementById("userById").addEventListener("click", (function () { _this.userById(); }));
                document.getElementById("userMessages").addEventListener("click", (function () { _this.loadUserMessages(); }));
                document.getElementById("userGroups").addEventListener("click", (function () { _this.loadUserGroups(); }));
                document.getElementById("userManager").addEventListener("click", (function () { _this.loadUserManager(); }));
                document.getElementById("groupsWithPaging").addEventListener("click", (function () { _this.loadGroupsWithPaging(); }));
                document.getElementById("groupById").addEventListener("click", (function () { _this.groupById(); }));
                document.getElementById("userPhoto").addEventListener("click", (function () { _this.loadUserPhoto(); }));
                document.getElementById("loggedIn").addEventListener("click", (function () { _this.isLoggedIn(); }));
                document.getElementById("whoAmI").addEventListener("click", (function () { _this.whoAmI(); }));
            });
        };
        //-----------------------------------------------Scenarios---------------------------------------------
        //Scenario 1: Logout
        App.prototype.logout = function () {
            this.identity.logOut();
        };
        //Scenario 2: Load users with paging
        App.prototype.loadUsersWithPaging = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            //this.graph.users(((users, error) => {
            //    this.getUsersCallback(users, null);
            //}), "$top=5");
            this.graph.usersAsync("$top=5").then((function (users) {
                _this.getUsersCallback(users, null);
            }));
        };
        //Scenario 3: Load users with custom odata query
        App.prototype.loadUsersWithOdataQuery = function (query) {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.usersAsync(query).then(function (users) { _this.getUsersCallback(users, null); });
        };
        //Scenario 4: Load user "me"
        App.prototype.loadUserMe = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(function (user) {
                document.getElementById("results").innerHTML += user.data.displayName + "</br>";
            });
        };
        //Scenario 5: Load user by ID
        App.prototype.userById = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.userAsync(document.getElementById("userId").value).then(function (user) {
                document.getElementById("results").innerHTML += user.data.displayName + "</br>";
            }).fail(function (error) {
                window.alert(error.status);
            });
        };
        //Scenario 6: Load user "me" and then its messages
        App.prototype.loadUserMessages = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.messagesAsync("$top=2").then(function (messages) {
                    _this.messagesCallback(messages, null);
                }).fail(function (error) {
                    _this.messagesCallback(null, error);
                });
            });
        };
        //Scenario 7: Load user "me" and then its groups
        App.prototype.loadUserGroups = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((function (user) {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                user.memberOf((function (groups, error) {
                    _this.groupsCallback(groups, error);
                }), "$top=5");
            }));
        };
        //Scenario 8: Load user "me" and then its manager
        App.prototype.loadUserManager = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                user.manager(function (manager) {
                    document.getElementById("results").innerHTML += manager.data.displayName + "</br>";
                });
            });
        };
        //Scenario 9: Load groups with paging
        App.prototype.loadGroupsWithPaging = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.groups((function (groups, error) {
                _this.getGroupsCallback(groups, null);
            }), "$top=5");
        };
        //Scenario 10: Load group by ID
        App.prototype.groupById = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.groupAsync(document.getElementById("groupId").value).then(function (group) {
                document.getElementById("results").innerHTML += group.data.displayName + "</br>";
            });
        };
        //Scenario 11: Load user "me" and then its messages
        App.prototype.loadUserPhoto = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";
                user.profilePhoto(function (photo, error) {
                    if (error)
                        window.alert(error.statusText);
                    else {
                        //Photo metadata
                        var x = photo;
                    }
                });
                user.profilePhotoValue(function (photoValue, error) {
                    if (error)
                        window.alert(error.statusText);
                    else {
                        var img = document.createElement("img");
                        var reader = new FileReader();
                        reader.onloadend = function () {
                            img.src = reader.result;
                        };
                        reader.readAsDataURL(photoValue);
                        document.getElementById("results").appendChild(img);
                    }
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
        App.prototype.getUsersCallback = function (users, error) {
            var _this = this;
            if (error) {
                document.getElementById("results").innerText = error.statusText;
                return;
            }
            users.data.forEach(function (item) {
                document.getElementById("results").innerHTML += item.data.displayName + "</br>";
            });
            if (users.nextLink) {
                users.nextLink().then((function (result) {
                    _this.getUsersCallback(result, null);
                }));
            }
        };
        App.prototype.getGroupsCallback = function (groups, error) {
            var _this = this;
            if (error) {
                document.getElementById("results").innerText = error.statusText;
                return;
            }
            groups.data.forEach(function (item) {
                document.getElementById("results").innerHTML += item.data.displayName + "</br>";
            });
            if (groups.nextLink) {
                groups.nextLink().then((function (result) {
                    _this.getGroupsCallback(result, null);
                }));
            }
        };
        App.prototype.messagesCallback = function (messages, error) {
            var _this = this;
            if (messages.data) {
                messages.data.forEach(function (item) {
                    document.getElementById("results").innerHTML += item.data.subject + "</br>";
                });
            }
            if (messages.nextLink) {
                messages.nextLink().then(function (messages) {
                    _this.messagesCallback(messages, error);
                });
            }
        };
        App.prototype.groupsCallback = function (groups, error) {
            var _this = this;
            groups.data.forEach(function (item) {
                document.getElementById("results").innerHTML += item.data.displayName + "</br>";
            });
            if (groups.nextLink) {
                groups.nextLink((function (groups, error) {
                    _this.groupsCallback(groups, error);
                }));
            }
        };
        return App;
    })();
    Sample.App = App;
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
