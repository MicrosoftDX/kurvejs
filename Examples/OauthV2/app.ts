
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Sample {
    class Error {
    }
    export class AppV2 {
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
                version: kurve.EndPointVersion.v2
            });

            //We can request for specific scopes during logon (so user will have to consent them right away and not during the flow of the app
            //The list of available consents is available under Kuve.Scopes module
            this.identity.loginAsync({ scopes: [kurve.Scopes.Mail.Read, kurve.Scopes.General.OpenId] }).then(() => {

                
                this.graph = new kurve.Graph({ identity: this.identity });
                
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
            this.identity.getAccessTokenForScopesAsync([kurve.Scopes.User.Read, kurve.Scopes.Mail.Read], false).then((token) => {
                document.getElementById("results").innerText = "token: " + token;
            });
        }
    
        //Scenario 3: Load user "me"
        private loadUserMe(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += user.item.displayName + "</br>";
            });
        }

        //Scenario 4: Load user "me" and then its messages
        private loadUserMessages(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.messages.GetMessages("$top=2").then((messages) => {

                    this.messagesCallback(messages,null );
                }).fail((error) =>{
                    this.messagesCallback(null,error);
                });
        }

        private loadUserEvents(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Messages:" + "</br>";
                user.self.events.GetEvents("$top=2").then((events)=>{
                    this.eventsCallback(events, null);
                }).fail((error) => {
                    this.eventsCallback(null, error);
                });

            });
        }


        //Scenario 5: Load user "me" and then its messages
        private loadUserPhoto(): void {
            document.getElementById("results").innerHTML = "";
            this.graph.me.GetUser().then((user) => {
                document.getElementById("results").innerHTML += "User:" + user.item.displayName + "</br>";
                document.getElementById("results").innerHTML += "Photo:" + "</br>";

                user.self.photo.GetPhotoProperties().then((photo) => {                  
                        var x = photo;
                });

                user.self.photo.GetPhotoImage().then((photoValue) => {
                    var img = document.createElement("img");
                    var reader = new FileReader();
                    reader.onloadend = () => {
                        img.src = reader.result;
                    }
                    reader.readAsDataURL(photoValue.item);

                    document.getElementById("results").appendChild(img);
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


        private messagesCallback(messages: kurve.Collection<kurve.MessageDataModel, kurve.Messages>, error: kurve.Error): void {
            if (messages.items) {
                messages.items.forEach((item) => {
                    document.getElementById("results").innerHTML += item.subject + "</br>";
                });
            }

            if (messages.next) {
                messages.next.GetMessages().then((messages) => {
                    this.messagesCallback(messages, error);
                });
            }
        }

        private eventsCallback(events: kurve.Collection<kurve.EventDataModel, kurve.Events>, error: kurve.Error): void {
            if (events.items) {
                events.items.forEach((event) => {
                    document.getElementById("results").innerHTML += event.subject + "</br>";
                });
            }

            if (events.next) {
                events.next.GetEvents().then((stuff) => {
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
