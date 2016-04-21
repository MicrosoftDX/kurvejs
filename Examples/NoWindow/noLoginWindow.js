// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Sample;
(function (Sample) {
    var Error = (function () {
        function Error() {
        }
        return Error;
    }());
    var AppNoWindow = (function () {
        //constructor() {
        //    this.test();
        //}
        function AppNoWindow(clientId, tokenProcessorUrl) {
            //Setup
            this.clientId = clientId;
            this.tokenProcessorUri = tokenProcessorUrl;
            //Create identity object
            this.identity = new kurve.Identity({ clientId: this.clientId, tokenProcessingUri: this.tokenProcessorUri, version: kurve.EndPointVersion.v1 });
            this.graph = new kurve.Graph({ identity: this.identity });
            if (this.identity.checkForIdentityRedirect()) {
                if (this.identity.isLoggedIn()) {
                    this.onLogin();
                }
            }
        }
        AppNoWindow.prototype.doLogin = function () {
            this.identity.loginNoWindowAsync().then(this.onLogin);
        };
        AppNoWindow.prototype.onLogin = function () {
            var _this = this;
            //Update UI
            document.getElementById("initDiv").style.display = "none";
            document.getElementById("loginDiv").style.display = "none";
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
        };
        //-----------------------------------------------Scenarios---------------------------------------------
        //Scenario 1: Logout
        AppNoWindow.prototype.logout = function () {
            this.identity.logOut();
        };
        //Scenario 2: Load users with paging
        AppNoWindow.prototype.loadUsersWithPaging = function () {
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
        AppNoWindow.prototype.loadUsersWithOdataQuery = function (query) {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.users.GetUsers(query).then(function (users) { _this.getUsersCallback(users, null); });
        };
        //Scenario 4: Load user "me"
        AppNoWindow.prototype.loadUserMe = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += user.data.displayName + "</br>";
            });
        };
        //Scenario 5: Load user by ID
        AppNoWindow.prototype.userById = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.users.$(document.getElementById("userId").value).GetUser().then(function (user) {
                document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            }).fail(function (error) {
                window.alert(error.status);
            });
        };
        //Scenario 6: Load user "me" and then its messages
        AppNoWindow.prototype.loadUserMessages = function () {
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
        //Scenario 7: Load user "me" and then its groups
        AppNoWindow.prototype.loadUserGroups = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                //TODO: Not supported right now
                //user.self.memberOf(((groups: Kurve.Groups, error: Kurve.Error) => {
                //    this.groupsCallback(groups, error);
                //}), "$top=5");
            }));
        };
        //Scenario 8: Load user "me" and then its manager
        AppNoWindow.prototype.loadUserManager = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                //TODO: Not supported yet
                //user.manager((manager: Kurve.User) => {
                //    document.getElementById("results").innerHTML += manager.data.displayName + "</br>";
                //});
            });
        };
        //Scenario 9: Load groups with paging
        AppNoWindow.prototype.loadGroupsWithPaging = function () {
            //Not supported yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.groups(((groups, error) => {
            //    this.getGroupsCallback(groups, null);
            //}), { top: "5" });
        };
        //Scenario 10: Load group by ID
        AppNoWindow.prototype.groupById = function () {
            //TODO: Not supported yet
            //document.getElementById("results").innerHTML = "";
            //this.graph.groupAsync((<HTMLInputElement>document.getElementById("groupId")).value).then((group) => {
            //    document.getElementById("results").innerHTML += group.data.displayName + "</br>";
            //});
        };
        //Scenario 11: Load user "me" and then its messages
        AppNoWindow.prototype.loadUserPhoto = function () {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(function (user) {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";
                //TODO: Not supported yet
                //user.profilePhoto((photo, error) => {
                //    if (error)
                //        window.alert(error.statusText);
                //    else {
                //        //Photo metadata
                //        var x = photo;
                //    }
                //});
                //user.profilePhotoValue((photoValue: any, error: Kurve.Error) => {
                //    if (error)
                //        window.alert(error.statusText);
                //    else {
                //        var img = document.createElement("img");
                //        var reader = new FileReader();
                //        reader.onloadend = () => {
                //            img.src = reader.result;
                //        }
                //        reader.readAsDataURL(photoValue);
                //        document.getElementById("results").appendChild(img);
                //    }
                //});
            });
        };
        //Scenario 12: Is logged in?
        AppNoWindow.prototype.isLoggedIn = function () {
            document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
        };
        //Scenario 13: Who am I?
        AppNoWindow.prototype.whoAmI = function () {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        };
        //--------------------------------Callbacks---------------------------------------------
        AppNoWindow.prototype.getUsersCallback = function (users, error) {
            var _this = this;
            if (error) {
                document.getElementById("results").innerText = error.statusText;
                return;
            }
            users.items.forEach(function (item) {
                document.getElementById("results").innerHTML += item.displayName + "</br>";
            });
            if (users.next) {
                users.next.GetUsers().then((function (result) {
                    _this.getUsersCallback(result, null);
                }));
            }
        };
        //TODO: Not supported yet
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
        AppNoWindow.prototype.messagesCallback = function (messages, error) {
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
        return AppNoWindow;
    }());
    Sample.AppNoWindow = AppNoWindow;
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
