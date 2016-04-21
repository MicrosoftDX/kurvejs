// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Sample;
(function (Sample) {
    var Error = (function () {
        function Error() {
        }
        return Error;
    }());
    var AppB2C = (function () {
        function AppB2C() {
            var _this = this;
            //Setup
            this.clientId = document.getElementById("AppID").value;
            this.redirectUri = document.getElementById("redirectUrl").value;
            //Create identity object
            this.identity = new kurve.Identity({ clientId: this.clientId, tokenProcessingUri: this.redirectUri, version: kurve.EndPointVersion.v2 });
            //We can request for specific scopes during logon (so user will have to consent them right away and not during the flow of the app
            //The list of available consents is available under Kuve.Scopes module
            this.identity.loginAsync({ policy: "B2C_1_facebooksignin", tenant: "matb2c.onmicrosoft.com" }).then(function () {
                //Update UI
                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (function () { _this.logout(); }));
                document.getElementById("loggedIn").addEventListener("click", (function () { _this.isLoggedIn(); }));
                document.getElementById("whoAmI").addEventListener("click", (function () { _this.whoAmI(); }));
            }).fail(function (error) {
                alert(JSON.stringify(error));
            });
        }
        //-----------------------------------------------Scenarios---------------------------------------------
        //Scenario 1: Logout
        AppB2C.prototype.logout = function () {
            this.identity.logOut();
        };
        //Scenario 2: Is logged in?
        AppB2C.prototype.isLoggedIn = function () {
            document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
        };
        //Scenario 3: Who am I?
        AppB2C.prototype.whoAmI = function () {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        };
        return AppB2C;
    }());
    Sample.AppB2C = AppB2C;
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
