// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Sample {
    export class App {
        private tenant;
        private clientId;
        private redirectUri;
        private identity: Kurve.Identity;
        private graph: Kurve.Graph;
        constructor() {

            //Setup
            this.tenant = (<HTMLInputElement>document.getElementById("tenant")).value;
            this.clientId = (<HTMLInputElement>document.getElementById("classID")).value;
            this.redirectUri = (<HTMLInputElement>document.getElementById("redirectUrl")).value;

            //Create identity object
            this.identity = new Kurve.Identity(this.tenant, this.clientId, this.redirectUri);


            //or  this.identity.loginAsync().then(() => {
            this.identity.login(() => {

                //////Option 1: Manualy passing the access token
                ////// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                ////this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                ////    this.graph = new Kurve.Graph(this.tenant, { defaultAccessToken: token });
                ////}));

                //Option 2: Automatically linking to the Identity object
                this.graph = new Kurve.Graph(this.tenant, { identity: this.identity });

                //Update UI

                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (() => { this.logout(); }));
                document.getElementById("usersWithPaging").addEventListener("click", (() => { this.loadUsersWithPaging(); }));
                document.getElementById("usersWithCustomODATA").addEventListener("click", (() => { this.loadUsersWithOdataQuery((<HTMLInputElement>document.getElementById('odataquery')).value); }));
                document.getElementById("meUser").addEventListener("click", (() => { this.loadUserMe(); }));
                document.getElementById("userMessages").addEventListener("click", (() => {
                    this.loadUserMessages();
                }));
                document.getElementById("loggedIn").addEventListener("click", (() => { this.isLoggedIn(); }));
                document.getElementById("whoAmI").addEventListener("click", (() => { this.whoAmI(); }));


            });
        }

        //-----------------------------------------------Scenarios---------------------------------------------

        //Scenario 1: Logout
        private logout(): void {
            this.identity.logOut();
        }

        //Scenario 2: Load users with paging
        private loadUsersWithPaging(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.users(((users, error) => {
                this.getUsersCallback(users, null);
            }), "$top=5");


            //this.graph.usersAsync("$top=5").then(((users) => {
            //    this.getUsersCallback(users, null);
            //}));
        }
    
        //Scenario 3: Load users with custom odata query
        private loadUsersWithOdataQuery(query: string): void {
            document.getElementById("results").innerHTML = "";
            this.graph.usersAsync(query).then((users) => { this.getUsersCallback(users, null); });
        }

        //Scenario 4: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user, error) => {
                if (error) {
                    document.getElementById("results").innerText = error;
                    return;
                }
                document.getElementById("results").innerHTML += user.displayName + "</br>";
            });
        }

        //Scenario 5: Load user "me" and then its messages
        private loadUserMessages(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.messages(((messages: any, error: string) => {
                    this.messagesCallback(messages, error);
                }), "$top=1");
            }));
        }

        //Scenario 6: Is logged in?
        private isLoggedIn(): void {
            document.getElementById("results").innerText = this.identity.isLoggedIn()?"True":"False";
        }

        //Scenario 7: Who am I?
        private whoAmI(): void {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        }
    

        //--------------------------------Callbacks---------------------------------------------

        //Callback used for scenario 1 to 3
        private getUsersCallback(users: any, error: string): void {
            if (error) {
                document.getElementById("results").innerText = error;
                return;
            }

            users.resultsPage.forEach((item) => {
                document.getElementById("results").innerHTML += item.displayName + "</br>";
            });

            if (users.nextLink) {
                users.nextLink().then(((result) => {
                    this.getUsersCallback(result, null);
                }));
            }
        }

        private messagesCallback(messages: any, error: string): void {
            messages.resultsPage.forEach((item) => {
                document.getElementById("results").innerHTML += item.subject + "</br>";
            });

            if (messages.nextLink) {
                messages.nextLink(((messages, error) => {
                    this.messagesCallback(messages, error);
                }));
            }

        }
    }
}

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
