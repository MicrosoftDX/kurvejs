// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Sample {
    class Error {
    }
    export class App {
        private clientId;
        private tokenProcessorUri;
        private identity: Kurve.Identity;
        private graph: Kurve.Graph;
        //constructor() {
        //    this.test();
        //}
        constructor() {
            //Setup
            this.clientId = (<HTMLInputElement>document.getElementById("classID")).value;
            this.tokenProcessorUri = (<HTMLInputElement>document.getElementById("tokenProcessorUrl")).value;

            //Create identity object
            this.identity = new Kurve.Identity(this.clientId, this.tokenProcessorUri);
            this.graph = new Kurve.Graph({ identity: this.identity });
            if (this.identity.checkForIdentityRedirect()) {
                if (this.identity.isLoggedIn()) { this.onLogin(); }
            }
        }
                
        public doLogin() : void {                                  
            this.identity.loginNoWindowAsync().then(this.onLogin);
        }
        
        public onLogin() 
        {                  
                //Update UI
                document.getElementById("initDiv").style.display = "none";
                document.getElementById("loginDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                
                document.getElementById("logoutBtn").addEventListener("click", (() => { this.logout(); }));
                document.getElementById("usersWithPaging").addEventListener("click", (() => { this.loadUsersWithPaging(); }));
                document.getElementById("usersWithCustomODATA").addEventListener("click", (() => { this.loadUsersWithOdataQuery((<HTMLInputElement>document.getElementById('odataquery')).value); }));
                document.getElementById("meUser").addEventListener("click", (() => { this.loadUserMe(); }));
                document.getElementById("userById").addEventListener("click", (() => { this.userById(); }));

                document.getElementById("userMessages").addEventListener("click", (() => {this.loadUserMessages();}));
                document.getElementById("userGroups").addEventListener("click", (() => {this.loadUserGroups();}));
                document.getElementById("userManager").addEventListener("click", (() => { this.loadUserManager(); }));
                document.getElementById("groupsWithPaging").addEventListener("click", (() => { this.loadGroupsWithPaging(); }));
                document.getElementById("groupById").addEventListener("click", (() => { this.groupById(); }));
                document.getElementById("userPhoto").addEventListener("click", (() => { this.loadUserPhoto(); }));
               
                document.getElementById("loggedIn").addEventListener("click", (() => { this.isLoggedIn(); }));
                document.getElementById("whoAmI").addEventListener("click", (() => { this.whoAmI(); }));
        }
        
        //-----------------------------------------------Scenarios---------------------------------------------
       

        //Scenario 1: Logout
        private logout(): void {
            this.identity.logOut();
        }

        //Scenario 2: Load users with paging
        private loadUsersWithPaging(): void {
            document.getElementById("results").innerHTML = "";

            //this.graph.users(((users, error) => {
            //    this.getUsersCallback(users, null);
            //}), "$top=5");


            this.graph.usersAsync("$top=5").then(((users) => {
                this.getUsersCallback(users, null);
            }));
        }
    
        //Scenario 3: Load users with custom odata query
        private loadUsersWithOdataQuery(query: string): void {
            document.getElementById("results").innerHTML = "";
            this.graph.usersAsync(query).then((users) => { this.getUsersCallback(users, null); });
        }

        //Scenario 4: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user) => {
                document.getElementById("results").innerHTML += user.data.displayName + "</br>";
            });
        }

        //Scenario 5: Load user by ID
        private userById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.userAsync((<HTMLInputElement>document.getElementById("userId")).value).then((user) => {
                document.getElementById("results").innerHTML += user.data.displayName + "</br>";
            }).fail((error) => {
                window.alert(error.status);
            });
        }

        //Scenario 6: Load user "me" and then its messages
        private loadUserMessages(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user: Kurve.User) => {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.messagesAsync("$top=2").then((messages: Kurve.Messages) => {
                    this.messagesCallback(messages,null );
                }).fail((error) =>{
                    this.messagesCallback(null,error);
                });

            });
        }

        //Scenario 7: Load user "me" and then its groups
        private loadUserGroups(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: Kurve.User) => {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                user.memberOf(((groups: Kurve.Groups, error: Kurve.Error) => {
                    this.groupsCallback(groups, error);
                }), "$top=5");
            }));
        }

        //Scenario 8: Load user "me" and then its manager
        private loadUserManager(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user: Kurve.User) => {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                user.manager((manager: Kurve.User) => {
                    document.getElementById("results").innerHTML += manager.data.displayName + "</br>";
                });
            });
        }

        //Scenario 9: Load groups with paging
        private loadGroupsWithPaging(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.groups(((groups, error) => {
                this.getGroupsCallback(groups, null);
            }), "$top=5");
        }

        //Scenario 10: Load group by ID
        private groupById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.groupAsync((<HTMLInputElement>document.getElementById("groupId")).value).then((group) => {
                document.getElementById("results").innerHTML += group.data.displayName + "</br>";
            });
        }

        //Scenario 11: Load user "me" and then its messages
        private loadUserPhoto(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me((user: Kurve.User) => {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";

                user.profilePhoto((photo, error) => {
                    if (error)
                        window.alert(error.statusText);
                    else {
                        //Photo metadata
                        var x = photo;
                    }
                });

                user.profilePhotoValue((photoValue: any, error: Kurve.Error) => {
                    if (error)
                        window.alert(error.statusText);
                    else {
                        var img = document.createElement("img");
                        var reader = new FileReader();
                        reader.onloadend = () => {
                            img.src = reader.result;
                        }
                        reader.readAsDataURL(photoValue);

                        document.getElementById("results").appendChild(img);
                    }
                });
            });
        }

        //Scenario 12: Is logged in?
        private isLoggedIn(): void {
            document.getElementById("results").innerText = this.identity.isLoggedIn()?"True":"False";
        }

        //Scenario 13: Who am I?
        private whoAmI(): void {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        }
    

        //--------------------------------Callbacks---------------------------------------------

        private getUsersCallback(users: Kurve.Users, error: Kurve.Error): void {
            if (error) {
                document.getElementById("results").innerText = error.statusText;
                return;
            }

            users.data.forEach((item) => {
                document.getElementById("results").innerHTML += item.data.displayName + "</br>";
            });

            if (users.nextLink) {
                users.nextLink().then(((result) => {
                    this.getUsersCallback(result, null);
                }));
            }
        }

        private getGroupsCallback(groups: Kurve.Groups, error: Kurve.Error): void {
            if (error) {
                document.getElementById("results").innerText = error.statusText;
                return;
            }

            groups.data.forEach((item) => {
                document.getElementById("results").innerHTML += item.data.displayName + "</br>";
            });

            if (groups.nextLink) {
                groups.nextLink().then(((result) => {
                    this.getGroupsCallback(result, null);
                }));
            }
        }

        private messagesCallback(messages: Kurve.Messages, error: Kurve.Error): void {
            if (messages.data) {
                messages.data.forEach((item) => {
                    document.getElementById("results").innerHTML += item.data.subject + "</br>";
                });
            }

            if (messages.nextLink) {
                messages.nextLink().then((messages) => {
                    this.messagesCallback(messages, error);
                });
            }

        }

        private groupsCallback(groups: Kurve.Groups, error: Kurve.Error): void {
            groups.data.forEach((item) => {
                document.getElementById("results").innerHTML += item.data.displayName + "</br>";
            });

            if (groups.nextLink) {
                groups.nextLink(((groups, error) => {
                    this.groupsCallback(groups, error);
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
