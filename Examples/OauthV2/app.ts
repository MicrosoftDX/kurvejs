// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Sample {
    class Error {
    }
    export class AppV2 {
        private clientId;
        private redirectUri;
        private identity: Kurve.Identity;
        private graph: Kurve.Graph;
       
        constructor() {
            //Setup
            this.clientId = (<HTMLInputElement>document.getElementById("AppID")).value;
            this.redirectUri = (<HTMLInputElement>document.getElementById("redirectUrl")).value;

            //Create identity object
            this.identity = new Kurve.Identity({
                clientId: this.clientId,
                tokenProcessingUri: this.redirectUri,
                version: Kurve.OAuthVersion.v2
            });

            //We can request for specific scopes during logon (so user will have to consent them right away and not during the flow of the app
            //The list of available consents is available under Kuve.Scopes module
            this.identity.loginAsync({ scopes: [Kurve.Scopes.Mail.Read, Kurve.Scopes.General.OpenId] }).then(() => {

                this.graph = new Kurve.Graph({ identity: this.identity });

                //Update UI

                document.getElementById("initDiv").style.display = "none";
                document.getElementById("scenarios").style.display = "";
                document.getElementById("logoutBtn").addEventListener("click", (() => { this.logout(); }));

                document.getElementById("requestAccessToken").addEventListener("click", (() => { this.requestAccessToken(); }));
                document.getElementById("meUser").addEventListener("click", (() => { this.loadUserMe(); }));              
                document.getElementById("userMessages").addEventListener("click", (() => {this.loadUserMessages();}));
                document.getElementById("userEvents").addEventListener("click", (() => { this.loadUserEvents(); }));
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

        //Scenario 2: Request access token
        private requestAccessToken(): void {
            document.getElementById("results").innerHTML = "";

            //Using false for prompt will cause it to attempt to silently get the token. If it fails, it will then re-attempt, this time
            //Prompting for a UI for consent
            this.identity.getAccessTokenForScopesAsync([Kurve.Scopes.User.Read, Kurve.Scopes.Mail.Read], false).then((token) => {
                document.getElementById("results").innerText = "token: " + token;
            });
        }
    
        //Scenario 3: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user) => {
                document.getElementById("results").innerHTML += user.data.displayName + "</br>";
            });
        }

        //Scenario 4: Load user "me" and then its messages
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

        private loadUserEvents(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.meAsync().then((user: Kurve.User) => {
                document.getElementById("results").innerHTML += "User:" + user.data.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.eventsAsync("$top=2").then((items: Kurve.Events) => {
                    this.eventsCallback(items, null);
                }).fail((error) => {
                    this.eventsCallback(null, error);
                });

            });
        }


        //Scenario 5: Load user "me" and then its messages
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

        //Scenario 6: Is logged in?
        private isLoggedIn(): void {
            document.getElementById("results").innerText = this.identity.isLoggedIn()?"True":"False";
        }

        //Scenario 7: Who am I?
        private whoAmI(): void {
            document.getElementById("results").innerText = JSON.stringify(this.identity.getIdToken());
        }
    

        //--------------------------------Callbacks---------------------------------------------


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

        private eventsCallback(events: Kurve.Events, error: Kurve.Error): void {
            if (events.data) {
                events.data.forEach((item) => {
                    document.getElementById("results").innerHTML += item.data.subject + "</br>";
                });
            }

            if (events.nextLink) {
                events.nextLink().then((stuff) => {
                    this.eventsCallback(stuff, error);
                });
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
