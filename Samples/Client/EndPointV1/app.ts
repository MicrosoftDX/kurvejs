// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Sample {
   
    export class AppV1 {
        private clientId;
        private redirectUri;
        private identity: kurve.Identity;
        private graph: kurve.Graph;
        constructor() {
            //Setup
            this.clientId = (<HTMLInputElement>document.getElementById("AppID")).value;
            this.redirectUri = (<HTMLInputElement>document.getElementById("redirectUrl")).value;

            //Create identity object
            this.identity = new kurve.Identity({
                clientId: this.clientId,
                tokenProcessingUri: this.redirectUri,
                version: kurve.EndPointVersion.v1
            });

            //or  this.identity.loginAsync().then(() => {
            this.identity.loginAsync().then(() => {

                ////Option 1: Manualy passing the access token
                //// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                //this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                //    this.graph = new Kurve.Graph({ defaultAccessToken: token });
                //}));

                //Option 2: Automatically linking to the Identity object
                this.graph = new kurve.Graph({ identity: this.identity });

                //Update UI

                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (() => { this.logout(); }));
                document.getElementById("usersWithPaging").addEventListener("click", (() => { this.loadUsersWithPaging(); }));
                document.getElementById("usersWithCustomODATA").addEventListener("click", (() => { this.loadUsersWithOdataQuery((<HTMLInputElement>document.getElementById('odataquery')).value); }));
                document.getElementById("meUser").addEventListener("click", (() => { this.loadUserMe(); }));
                document.getElementById("userById").addEventListener("click", (() => { this.userById(); }));

                document.getElementById("userMessages").addEventListener("click", (() => { this.loadUserMessages(); }));
                document.getElementById("userEvents").addEventListener("click", (() => { this.loadUserEvents(); }));
                document.getElementById("userGroups").addEventListener("click", (() => { this.loadUserGroups(); }));
                document.getElementById("userManager").addEventListener("click", (() => { this.loadUserManager(); }));
                document.getElementById("groupsWithPaging").addEventListener("click", (() => { this.loadGroupsWithPaging(); }));
                document.getElementById("groupById").addEventListener("click", (() => { this.groupById(); }));
                document.getElementById("userPhoto").addEventListener("click", (() => { this.loadUserPhoto(); }));

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
            this.showUsers(this.graph.users, "$top=5");
        }

        //Scenario 3: Load users with custom odata query
        private loadUsersWithOdataQuery(query: string): void {
            document.getElementById("results").innerHTML = "";
            this.showUsers(this.graph.users, query);
        }

        //Scenario 4: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user =>
                document.getElementById("results").innerHTML += user.item.displayName + "</br>"
            );
        }

        //Scenario 5: Load user by ID
        private userById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.users.$((<HTMLInputElement>document.getElementById("userId")).value).GetUser().then((user) => {
                document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            }).fail((error) => {
                window.alert(error.status);
            });
        }

        //Scenario 6: Load user "me" and then its messages
        private loadUserMessages(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                this.showMessages(user.self.messages, "$top=2");
            });
        }


        //Scenario 6: Load user "me" and then its events
        private loadUserEvents(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Events:" + "</br>";
                this.showEvents(user.self.events, "$top=2");
            });
        }

        //Scenario 7: Load user "me" and then its groups
        private loadUserGroups(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                this.showGroups(user.self.memberOf, "$top=5");
            });
        }

        //Scenario 8: Load user "me" and then its manager
        private loadUserManager(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                user.self.manager.GetManager().then((manager) => {
                    document.getElementById("results").innerHTML += manager.item.displayName + "</br>";
                });
            });
        }

        //Scenario 9: Load groups with paging
        private loadGroupsWithPaging(): void {
            document.getElementById("results").innerHTML = "";
            this.showGroups(this.graph.groups, "$top=5");
        }

        //Scenario 10: Load group by ID
        private groupById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.groups.$((<HTMLInputElement>document.getElementById("groupId")).value).GetGroup().then((group) => {
                document.getElementById("results").innerHTML += group.item.displayName + "</br>";
            });
        }

        //Scenario 11: Load user "me" and then its messages
        private loadUserPhoto(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";

                user.self.photo.GetPhotoProperties().then((photo) => {
                        //Photo metadata
                        var x = photo;
                });

                user.self.photo.GetPhotoImage().then((photoValue) => {
                    var img = document.createElement("img");
                    var reader = new FileReader();
                    reader.onloadend = () => {
                        img.src = reader.result;
                    }
                    reader.readAsDataURL(photoValue.raw);

                    document.getElementById("results").appendChild(img);
                });
            });
        }

        //Scenario 12: Is logged in?
        private isLoggedIn(): void {
            document.getElementById("results").innerText = this.identity.isLoggedIn() ? "True" : "False";
        }

        //Scenario 13: Who am I?
        private whoAmI(): void {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        }


        //--------------------------------Callbacks---------------------------------------------

        private showUsers(users: kurve.Users, odataQuery?: string): void {
            users.GetUsers(odataQuery)
                .then(collection => {
                    collection.items.forEach(user =>
                        document.getElementById("results").innerHTML += user.displayName + "</br>"
                    );

                    if (collection.next)
                        this.showUsers(collection.next);
                })
                .fail(error =>
                    document.getElementById("results").innerText = error.statusText
                )
        }


        private showGroups(groups: kurve.Groups, odataQuery?: string): void {
            groups.GetGroups(odataQuery)
                .then(collection => {
                    collection.items.forEach(group =>
                        document.getElementById("results").innerHTML += group.displayName + "</br>"
                    );

                    if (collection.next)
                        this.showGroups(collection.next);
                })
                .fail((error) => {
                    document.getElementById("results").innerText = JSON.stringify(error);
                });
        }

        private showMessages(messages: kurve.Messages, odataQuery?: string): void {
            messages.GetMessages(odataQuery)
                .then(collection => {
                    collection.items.forEach(message =>
                        document.getElementById("results").innerHTML += message.subject + "</br>"
                    );

                    if (collection.next)
                        this.showMessages(collection.next);
                })
                .fail(error =>
                    document.getElementById("results").innerText = error.statusText
                )

        }
       

        private showEvents(events: kurve.Events, odataQuery?: string): void {
            events.GetEvents(odataQuery)
                .then(collection => {
                    collection.items.forEach(event =>
                        document.getElementById("results").innerHTML += event.subject + "</br>"
                    );

                    if (events.next)
                        this.showEvents(collection.next);
                })
                .fail(error =>
                    document.getElementById("results").innerText = error.statusText
                )

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
