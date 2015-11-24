// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Sample;
(function (Sample) {
    var App = (function () {
        function App() {
            var _this = this;
            //Setup
            this.tenant = document.getElementById("tenant").value;
            this.clientId = document.getElementById("classID").value;
            this.redirectUri = document.getElementById("redirectUrl").value;
            //Create identity object
            this.identity = new Kurve.Identity(this.tenant, this.clientId, this.redirectUri);
            //or  this.identity.loginAsync().then(() => {
            this.identity.login(function () {
                //////Option 1: Manualy passing the access token
                ////// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                ////this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                ////    this.graph = new Kurve.Graph(this.tenant, { defaultAccessToken: token });
                ////}));
                //Option 2: Automatically linking to the Identity object
                _this.graph = new Kurve.Graph(_this.tenant, { identity: _this.identity });
                //Update UI
                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (function () { _this.logout(); }));
                document.getElementById("usersWithPaging").addEventListener("click", (function () { _this.loadUsersWithPaging(); }));
                document.getElementById("usersWithCustomODATA").addEventListener("click", (function () { _this.loadUsersWithOdataQuery(document.getElementById('odataquery').value); }));
                document.getElementById("meUser").addEventListener("click", (function () { _this.loadUserMe(); }));
                document.getElementById("userMessages").addEventListener("click", (function () {
                    _this.loadUserMessages();
                }));
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
            this.graph.users((function (users, error) {
                _this.getUsersCallback(users, null);
            }), "$top=5");
            //this.graph.usersAsync("$top=5").then(((users) => {
            //    this.getUsersCallback(users, null);
            //}));
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
            this.graph.meAsync().then(function (user, error) {
                if (error) {
                    document.getElementById("results").innerText = error;
                    return;
                }
                document.getElementById("results").innerHTML += user.displayName + "</br>";
            });
        };
        //Scenario 5: Load user "me" and then its messages
        App.prototype.loadUserMessages = function () {
            var _this = this;
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((function (user) {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.messages((function (messages, error) {
                    _this.messagesCallback(messages, error);
                }), "$top=1");
            }));
        };
        //Scenario 6: Is logged in?
        App.prototype.isLoggedIn = function () {
            document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
        };
        //Scenario 7: Who am I?
        App.prototype.whoAmI = function () {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        };
        //--------------------------------Callbacks---------------------------------------------
        //Callback used for scenario 1 to 3
        App.prototype.getUsersCallback = function (users, error) {
            var _this = this;
            if (error) {
                document.getElementById("results").innerText = error;
                return;
            }
            users.resultsPage.forEach(function (item) {
                document.getElementById("results").innerHTML += item.displayName + "</br>";
            });
            if (users.nextLink) {
                users.nextLink().then((function (result) {
                    _this.getUsersCallback(result, null);
                }));
            }
        };
        App.prototype.messagesCallback = function (messages, error) {
            var _this = this;
            messages.resultsPage.forEach(function (item) {
                document.getElementById("results").innerHTML += item.subject + "</br>";
            });
            if (messages.nextLink) {
                messages.nextLink((function (messages, error) {
                    _this.messagesCallback(messages, error);
                }));
            }
        };
        return App;
    })();
    Sample.App = App;
})(Sample || (Sample = {}));
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
