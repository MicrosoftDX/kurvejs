/// <reference path="../../../dist/kurve-global.d.ts" />
const kurve = window["Kurve"] as typeof Kurve;

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
    const init = () => new App();

    class App {
        private identity: Kurve.Identity;
        private graph: Kurve.Graph;

        constructor() {
            // Setup
            const endpointVersion = (<HTMLSelectElement>document.getElementById("model")).value == 'v2' ? kurve.EndpointVersion.v2 : kurve.EndpointVersion.v1;
            const clientId = (<HTMLInputElement>document.getElementById("AppID")).value || endpointVersion == kurve.EndpointVersion.v2 ? "13c5e4af-5ea6-4b48-8989-ca25c96ba1c4" : "636e98ea-3024-4810-a66e-cda4bfa0a489";
            const loc = document.URL;
            const redirectUri = loc.substr(0, loc.indexOf("/Samples/Client/VanillaJS")) + "/dist/login.html";

            // Create identity object
            this.identity = new kurve.Identity(clientId, redirectUri, { endpointVersion: endpointVersion}); 
            
            const scopes = endpointVersion == kurve.EndpointVersion.v2 ? { scopes: [kurve.Scopes.Mail.Read, kurve.Scopes.General.OpenId] } : {};

            // Login
            this.identity.loginAsync(scopes).then(_ => {

                ////Option 1: Manualy passing the access token
                //// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                //this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                //    this.graph = new Kurve.Graph({ defaultAccessToken: token });
                //}));

                //Option 2: Automatically linking to the Identity object
                this.graph = new kurve.Graph(this.identity);

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
            this.graph.users.GetUsers().then(users =>
                this.showUsers(users)
            );
        }

        //Scenario 3: Load users with custom odata query
        private loadUsersWithOdataQuery(query: string): void {
            document.getElementById("results").innerHTML = "";
            this.graph.users.GetUsers(query).then(users =>
                this.showUsers(users)
            );
        }

        //Scenario 4: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user =>
                document.getElementById("results").innerHTML += user.displayName + "</br>"
            ).fail(error =>
                document.getElementById("results").innerText = JSON.stringify(error)
            );
        }

        //Scenario 5: Load user by ID
        private userById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.users.$((<HTMLInputElement>document.getElementById("userId")).value).GetUser().then((user) => {
                document.getElementById("results").innerHTML += user.displayName + "</br>";
            }).fail((error) => {
                window.alert(error.status);
            });
        }

        //Scenario 6: Load user "me" and then its messages
        private loadUserMessages(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user._context.messages.GetMessages().then(messages =>
                    this.showMessages(messages)
                );
            });
        }


        //Scenario 6: Load user "me" and then its events
        private loadUserEvents(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Events:" + "</br>";
                user._context.events.GetEvents().then(events =>
                    this.showEvents(events)
                );
            });
        }

        //Scenario 7: Load user "me" and then its groups
        private loadUserGroups(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then(user => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                user._context.memberOf.GetGroups("$top=5").then(groups =>
                    this.showGroups(groups)
                );
            });
        }

        //Scenario 8: Load user "me" and then its manager
        private loadUserManager(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                user._context.manager.GetUser().then(manager => {
                    document.getElementById("results").innerHTML += manager.displayName + "</br>";
                });
            });
        }

        //Scenario 9: Load groups with paging
        private loadGroupsWithPaging(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.groups.GetGroups().then(groups =>
                this.showGroups(groups, 5)
            );
        }

        //Scenario 10: Load group by ID
        private groupById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.groups.$((<HTMLInputElement>document.getElementById("groupId")).value).GetGroup().then((group) => {
                document.getElementById("results").innerHTML += group.displayName + "</br>";
            });
        }

        //Scenario 11: Load user "me" and then its messages
        private loadUserPhoto(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";

                user._context.photo.GetPhotoProperties().then((photo) => {
                        //Photo metadata
                        var x = photo;
                });

                user._context.photo.GetPhotoImage().then((photoValue) => {
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

        private showUsers(users: Kurve.GraphCollection<Kurve.UserDataModel, Kurve.Users, Kurve.User>, limit:number = 5): void {
            for (let user of users.value) {
                document.getElementById("results").innerHTML += user.displayName + "</br>";
                if (limit-- <= 0)
                    return;
            }

            users._next && users._next().then(nextUsers =>
                this.showUsers(nextUsers, limit)
            ).fail(error =>
                document.getElementById("results").innerText = JSON.stringify(error)
            );
        }


        private showGroups<GraphNode extends Kurve.CollectionNode>(groups: Kurve.GraphCollection<Kurve.GroupDataModel, GraphNode, Kurve.Group>, limit:number = 5): void {
            for (let group of groups.value) {
                document.getElementById("results").innerHTML += group.displayName + "</br>"
                if (limit-- <= 0)
                    return;
            }

            groups._next && groups._next().then(nextGroups =>
                this.showGroups(nextGroups)
            ).fail(error =>
                document.getElementById("results").innerText = JSON.stringify(error)
            );
        }

        private showMessages(messages: Kurve.GraphCollection<Kurve.MessageDataModel, Kurve.Messages, Kurve.Message>, limit:number = 5): void {
            for (let message of messages.value) {
                document.getElementById("results").innerHTML += message.subject + "</br>"
                if (limit-- <= 0)
                    return;
            }
            
            messages._next && messages._next().then(nextMessages =>
                this.showMessages(nextMessages, limit)
            ).fail(error =>
                document.getElementById("results").innerText = error.statusText
            );
        }
       

        private showEvents(events: Kurve.GraphCollection<Kurve.EventDataModel, Kurve.Events, Kurve.Event>, limit:number = 5): void {
            for (let event of events.value) {
                document.getElementById("results").innerHTML += event.subject + "</br>"
                if (limit-- <= 0)
                    return;
            }

            events._next && events._next().then(nextEvents =>
                this.showEvents(nextEvents)
            ).fail(error =>
                document.getElementById("results").innerText = error.statusText
            );
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
