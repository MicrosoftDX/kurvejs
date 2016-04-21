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
            var _this = this;
            document.getElementById("results").innerHTML = "";
            //this.graph.users(((users, error) => {
            //    this.getUsersCallback(users, null);
            //}), "$top=5");
            this.graph.users.GetUsers("$top=5").then((function (users) {
                _this.getUsersCallback(users, null);
            }));
        };
        //Scenario 3: Load users with custom odata query
        AppV1.prototype.loadUsersWithOdataQuery = function (query) {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.users.GetUsers(query).then(function (users) { _this.getUsersCallback(users, null); });
        };
        //Scenario 4: Load user "me"
        AppV1.prototype.loadUserMe = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            });
        };
        //Scenario 5: Load user by ID
        AppV1.prototype.userById = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.users.$(document.getElementById("userId").value).GetUser().then(function (user) {
                document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            }).fail(function (error) {
                window.alert(error.status);
            });
        };
        //Scenario 6: Load user "me" and then its messages
        AppV1.prototype.loadUserMessages = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.self.messages.GetMessages("$top=2").then(function (messages) {
                    _this.messagesCallback(messages, null);
                }).fail(function (error) {
                    _this.messagesCallback(null, error);
                });
            });
        };
        //Scenario 6: Load user "me" and then its messages
        AppV1.prototype.loadUserEvents = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Events:" + "</br>";
                user.self.events.GetEvents("$top=2").then(function (items) {
                    _this.eventsCallback(items, null);
                }).fail(function (error) {
                    _this.eventsCallback(null, error);
                });
            });
        };
        //Scenario 7: Load user "me" and then its groups
        AppV1.prototype.loadUserGroups = function () {
            //TODO: Not implemented yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.me.GetUser().then(((user) => {
            //    document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
            //    document.getElementById("results").innerHTML += "Groups:" + "</br>";
            //    user.self.memberOf(((groups, error) => {
            //        this.groupsCallback(groups, error);
            //    }), "$top=5");
            //}));
        };
        //Scenario 8: Load user "me" and then its manager
        AppV1.prototype.loadUserManager = function () {
            //TODO: not implemented yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.me.GetUser().then((user) => {
            //    document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
            //    document.getElementById("results").innerHTML += "Manager:" + "</br>";
            //    user.self.manager((manager: Kurve.User) => {
            //        document.getElementById("results").innerHTML += manager.data.displayName + "</br>";
            //    });
            //});
        };
        //Scenario 9: Load groups with paging
        AppV1.prototype.loadGroupsWithPaging = function () {
            //TODO: not implemented yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.groups(((groups, error) => {
            //    this.getGroupsCallback(groups, null);
            //}), "$top=5");
        };
        //Scenario 10: Load group by ID
        AppV1.prototype.groupById = function () {
            //Not implemented yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.groupAsync((<HTMLInputElement>document.getElementById("groupId")).value).then((group) => {
            //    document.getElementById("results").innerHTML += group.data.displayName + "</br>";
            //});
        };
        //Scenario 11: Load user "me" and then its messages
        AppV1.prototype.loadUserPhoto = function () {
            //Not implemented yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.me((user: Kurve.User) => {
            //    document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
            //    document.getElementById("results").innerHTML += "Photo:" + "</br>";
            //    user.profilePhoto((photo, error) => {
            //        if (error)
            //            window.alert(error.statusText);
            //        else {
            //            //Photo metadata
            //            var x = photo;
            //        }
            //    });
            //    user.profilePhotoValue((photoValue: any, error: Kurve.Error) => {
            //        if (error)
            //            window.alert(error.statusText);
            //        else {
            //            var img = document.createElement("img");
            //            var reader = new FileReader();
            //            reader.onloadend = () => {
            //                img.src = reader.result;
            //            }
            //            reader.readAsDataURL(photoValue);
            //            document.getElementById("results").appendChild(img);
            //        }
            //    });
            //});
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
        AppV1.prototype.getUsersCallback = function (users, error) {
            var _this = this;
            if (error) {
                document.getElementById("results").innerText = error.statusText;
                return;
            }
            users.items.forEach(function (user) {
                document.getElementById("results").innerHTML += user.displayName + "</br>";
            });
            if (users.next) {
                //TODO: Bug, requires Odata
                users.next.GetUsers("").then((function (result) {
                    _this.getUsersCallback(result, null);
                }));
            }
        };
        //TODO: Missing groups implementation
        //private getGroupsCallback(groups: kurve.Collection<kurve.GroupDataModel, kurve.Groups>, error: kurve.Error): void {
        //    if (error) {
        //        document.getElementById("results").innerText = error.statusText;
        //        return;
        //    }
        //    groups.data.forEach((item) => {
        //        document.getElementById("results").innerHTML += item.data.displayName + "</br>";
        //    });
        //    if (groups.nextLink) {
        //        groups.nextLink().then(((result) => {
        //            this.getGroupsCallback(result, null);
        //        }));
        //    }
        //}
        AppV1.prototype.messagesCallback = function (messages, error) {
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
        AppV1.prototype.eventsCallback = function (events, error) {
            var _this = this;
            if (events.items) {
                events.items.forEach(function (event) {
                    document.getElementById("results").innerHTML += event.subject + "</br>";
                });
            }
            if (events.next) {
                events.next.GetEvents().then(function (results) {
                    _this.eventsCallback(results, error);
                });
            }
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
