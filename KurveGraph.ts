// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {

    export class Graph {
        private req: XMLHttpRequest = null;
        private state: string = null;
        private nonce: string = null;
        private accessToken: string = null;
        private tenantId: string = null;
        private KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/beta/";

        constructor(tenantId: string, identityInfo: { identity: Identity });
        constructor(tenantId: string, identityInfo: { defaultAccessToken: string });
        constructor(tenantId: string, identityInfo: any) {
            this.tenantId = tenantId;
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }

        //Messages
      
        //Users
    
        public meAsync(odataQuery?: string): Promise {
            var d = new Deferred();
            this.me((user, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(user);
                }
            }, odataQuery);
            return d.promise;
        }

        public me(callback: (user: any, error: string) => void, odataQuery?: string): void {
            var urlString: string = this.buildMeUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUser(urlString, callback);
        }

        public usersAsync(odataQuery?: string): Promise {
            var d = new Deferred();
            this.users((users, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(users);
                }
            }, odataQuery);
            return d.promise;
        }

        public users(callback: (users: any, error: string) => void, odataQuery?: string): void {
            var urlString: string = this.buildUsersUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUsers(urlString, callback);
        }

        //http verbs

        public getAsync(url: string): Promise {
            var d = new Deferred();
            this.get(url, (response) => {
                d.resolve(response);
            });
            return d.promise;
        }

        public get(url: string, callback: (response:string)=>void): void {
            var xhr = new XMLHttpRequest();

            xhr.onreadystatechange = (() => {
                if (xhr.readyState === 4) {
                    callback(xhr.responseText);
                }
            });

            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr);
        }


        //Private methods

        private getUsers(urlString, callback: (users: any, error: string) => void): void {
            this.get(urlString, ((result: string) => {

                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    callback(null, JSON.stringify(usersODATA.error));
                    return;
                }

                var resultsArray = !usersODATA.value ? [usersODATA] : usersODATA.value;

                for (var i: number = 0; i < resultsArray.length; i++) {
                    this.decorateUserObject(resultsArray[i]);
                }

                var users = {
                    resultsPage: resultsArray
                };

                //implement nextLink
                var nextLink = usersODATA['@odata.nextLink'];

                if (nextLink) {
                    (<any>users).nextLink = ((callback?: (result: string, error: string) => void) => {
                        var d = new Deferred();
                        this.getUsers(nextLink, ((result, error) => {
                            if (callback)
                                callback(result, error);
                            else if (error)
                                d.reject(error);
                            else
                                d.resolve(result);
                        }));
                        return d.promise;
                    });
                }

                callback(users, null);
            }));
        }

        private getUser(urlString, callback: (user: any, error: string) => void): void {
            this.get(urlString, (result: string) => {

                var userODATA = JSON.parse(result);
                if (userODATA.error) {
                    callback(null, JSON.stringify(userODATA.error));
                    return;
                }

                this.decorateUserObject(userODATA);

                callback(userODATA, null);
            });

        }

        private addAccessTokenAndSend(xhr: XMLHttpRequest): void {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            } else {
                //Using the integrated Identity object
                this.KurveIdentity.getAccessToken(this.defaultResourceID, ((token: string, error: string) => {
                    //cache the token
                    if (error)
                        throw error;
                    xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                    xhr.send();
                }));
            }
        }

        private decorateUserObject(user: any): void {
            user.messages = ((callback: (messages: any, error: string) => void, odataQuery?: string) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/messages";
                if (odataQuery) urlString += "?" + odataQuery;
                this.getMessages(urlString, callback, odataQuery);
            });
        }

        private decorateMessageObject(message: any): void {
        }

        private getMessages(urlString: string, callback: (messages: any, error: string) => void, odataQuery?: string): void {

            var url = urlString;
            if (odataQuery) urlString += "?" + odataQuery;
            this.get(url, ((result: string) => {

                var messagesODATA = JSON.parse(result);
                if (messagesODATA.error) {
                    callback(null, JSON.stringify(messagesODATA.error));
                    return;
                }

                var resultsArray = !messagesODATA.value ? [messagesODATA] : messagesODATA.value;

                for (var i: number = 0; i < resultsArray.length; i++) {
                    this.decorateMessageObject(resultsArray[i]);
                }

                var messages = {
                    resultsPage: resultsArray
                };
                var nextLink = messagesODATA['@odata.nextLink'];
                //implement nextLink
                if (nextLink) {
                    (<any>messages).nextLink = ((callback?: (result: string, error: string) => void) => {
                        var d = new Deferred();
                        this.getMessages(nextLink, ((result, error) => {
                            if (callback)
                                callback(result, error);
                            else if (error)
                                d.reject(error);
                            else
                                d.resolve(result);
                        }));
                        return d.promise;
                    });
                }

                callback(messages, null);
            }));
        }       

        private buildMeUrl(): string {
            return this.baseUrl + "me";
        }
        private buildUsersUrl(): string {
            return this.baseUrl + this.tenantId + "/users";
        }
    }
}


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
