// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Sample {
    export class App {
        private clientId : string;
        private redirectUri : string;
        private identity: Kurve.Identity;
        private graph: Kurve.Graph;
        constructor() {


            //Setup
            this.clientId = (<HTMLInputElement>document.getElementById("classID")).value;
            this.redirectUri = (<HTMLInputElement>document.getElementById("redirectUrl")).value;

            //Create identity object
            this.identity = new Kurve.Identity(this.clientId, this.redirectUri);


            //or  this.identity.loginAsync().then(() => {
            this.identity.login(() => {

                //////Option 1: Manualy passing the access token
                ////// or... this.identity.getAccessToken("https://graph.microsoft.com", ((token) => {
                ////this.identity.getAccessTokenAsync("https://graph.microsoft.com").then(((token) => {
                ////    this.graph = new Kurve.Graph({ defaultAccessToken: token });
                ////}));


                //Option 2: Automatically linking to the Identity object
                this.graph = new Kurve.Graph({ identity: this.identity });

                //Update UI
                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (() => { this.logout(); }));

                // Login status
                document.getElementById("loggedIn").addEventListener("click", (() => { this.isLoggedIn(); }));
                document.getElementById("whoAmI").addEventListener("click", (() => { this.whoAmI(); }));

                // Me information
                document.getElementById("meUser").addEventListener("click", (() => { this.loadUserMe(); }));
                document.getElementById("userPhoto").addEventListener("click", (() => { this.loadUserPhoto(); }));
                document.getElementById("userMessages").addEventListener("click", (() => {this.loadUserMessages();}));
                document.getElementById("userCalendar").addEventListener("click", (() => {this.loadUserCalendar();}));
                document.getElementById("userContacts").addEventListener("click", (() => {this.loadUserContacts();}));
                document.getElementById("userGroups").addEventListener("click", (() => {this.loadUserGroups();}));
                document.getElementById("userManager").addEventListener("click", (() => { this.loadUserManager(); }));

                // Global information
                document.getElementById("userById").addEventListener("click", (() => { this.userById(); }));
                document.getElementById("usersWithPaging").addEventListener("click", (() => { this.loadUsersWithPaging(); }));
                document.getElementById("usersWithCustomODATA").addEventListener("click", (() => { this.loadUsersWithOdataQuery((<HTMLInputElement>document.getElementById('odataquery')).value); }));
                document.getElementById("groupsWithPaging").addEventListener("click", (() => { this.loadGroupsWithPaging(); }));
                document.getElementById("groupById").addEventListener("click", (() => { this.groupById(); }));
            });
        }

        //-----------------------------------------------Scenarios---------------------------------------------

        
        private logout(): void {
            this.identity.logOut();
        }

        private isLoggedIn(): void {
            document.getElementById("results").innerText = this.identity.isLoggedIn()?"True":"False";
        }

        private whoAmI(): void {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        }

        //Scenario 4: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user) => {
                document.getElementById("results").innerHTML += user.displayName + "</br>";
            });
        }

        //Scenario 5: Load user by ID
        private userById(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.userAsync((<HTMLInputElement>document.getElementById("userId")).value).then((user) => {
                document.getElementById("results").innerHTML += user.displayName + "</br>";
            });
        }


        private loadUserCalendar(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Calender:" + "</br>";
                user.CalendarAsync("$top=2").then((calendarEntries: any) => {
                    this.calendarCallback(calendarEntries, null);
                }).fail((error) => { this.calendarCallback(null, error); });

                //or...
                //user.messagesAsync("$top=2").then((messages: any) => {
                //    this.messagesCallback(messages, null);
                //});
            }));
        }

        private loadUserContacts(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Contacts:" + "</br>";
                user.ContactsAsync("$top=2").then((contacts: any) => {
                    this.contactsCallback(contacts, null);
                }).fail((error) => { this.calendarCallback(null, error); });
            }));
        }
        
        private loadUserMessages(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.messages(((messages: any, error: string) => {
                    if (error) { messages = null; }                   
                    this.messagesCallback(messages, error);
                }), "$top=2");
            }));
        }


        // Global directory information
        private loadUserGroups(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Groups:" + "</br>";
                user.memberOf(((groups: any, error: string) => {
                    this.groupsCallback(groups, error);
                }), "$top=5");
            }));
        }

        //Scenario 8: Load user "me" and then its manager
        private loadUserManager(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then(((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Manager:" + "</br>";
                user.managerAsync().then((manager: any) => {
                    document.getElementById("results").innerHTML += manager.displayName + "</br>";
                });
            }));
        }


        private loadUserPhoto(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user: any) => {
                document.getElementById("results").innerHTML += "User:" + user.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";
                user.photoValue(((photoValue: any, error: string) => {
                    if (error)
                        window.alert(error);
                    else {
                        var img = document.createElement("img");
                        var reader = new FileReader();
                        reader.onloadend = () => {
                            img.src = reader.result;
                        }
                        reader.readAsDataURL(photoValue);

                        document.getElementById("results").appendChild(img);
                    }
                }));
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
                document.getElementById("results").innerHTML += group.displayName + "</br>";
            });
        }

        private loadUsersWithPaging(): void {
            document.getElementById("results").innerHTML = "";

            this.graph.users(((users, error) => {
                this.getUsersCallback(users, null);
            }), "$top=5");


            //this.graph.usersAsync("$top=5").then(((users) => {
            //    this.getUsersCallback(users, null);
            //}));
        }
    
        private loadUsersWithOdataQuery(query: string): void {
            document.getElementById("results").innerHTML = "";
            this.graph.usersAsync(query).then((users) => { this.getUsersCallback(users, null); });
        }
    

        //--------------------------------Callbacks---------------------------------------------

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

        private getGroupsCallback(groups: any, error: string): void {
            if (error) {
                document.getElementById("results").innerText = error;
                return;
            }

            groups.resultsPage.forEach((item) => {
                document.getElementById("results").innerHTML += item.displayName + "</br>";
            });

            if (groups.nextLink) {
                groups.nextLink().then(((result) => {
                    this.getGroupsCallback(result, null);
                }));
            }
        }

        private calendarCallback(calendarEntries: any, error: string): void {
            calendarEntries.resultsPage.forEach((item) => {
                document.getElementById("results").innerHTML += item.subject + "</br>";
            });

            if (calendarEntries.nextLink) {
                calendarEntries.nextLink(((messages, error) => {
                    this.messagesCallback(messages, error);
                }));
            }

        }
        
        private contactsCallback(contacts: any, error: string): void {
            contacts.resultsPage.forEach((item) => {
                document.getElementById("results").innerHTML += item.subject + "</br>";
            });

            if (contacts.nextLink) {
                contacts.nextLink((c, e) => {
                    this.messagesCallback(c, e);
                });
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

        private groupsCallback(groups: any, error: string): void {
            groups.resultsPage.forEach((item) => {
                document.getElementById("results").innerHTML += item.displayName + "</br>";
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
