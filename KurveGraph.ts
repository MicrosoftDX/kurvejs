// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {

    export class User {
        public graph: Kurve.Graph;
        public businessPhones = "";
        public displayName: string = "";
        public givenName: string;
        public jobTitle: string;
        public mail: string;
        public mobilePhone: string;
        public officeLocation: string;
        public preferredLanguage: string;
        public surname: string;
        public userPrincipalName: string = "";
        public id: string;

        constructor(graph: Kurve.Graph, o: { userPrincipalName: string, displayName : string }) {
            this.graph = graph;
            this.userPrincipalName = o.userPrincipalName;
            this.displayName = o.displayName;
        }

        // These are all passthroughs to the graph

        public memberOf(callback: (groups: Group[], Error) => void, Error, odataQuery?: string) {
            this.graph.memberOfForUser(this.userPrincipalName, callback, odataQuery);
        }

        public memberOfAsync(odataQuery?: string): Promise {
            return this.graph.memberOfForUserAsync(this.userPrincipalName, odataQuery);
        }

        public messages(callback: (messages: Message[], error: Error) => void, odataQuery?: string) {
            this.graph.messagesForUser(this.userPrincipalName, callback, odataQuery);
        }

        public messagesAsync(odataQuery?: string): Promise {
            return this.graph.messagesForUserAsync(this.userPrincipalName, odataQuery);
        }

        public manager(callback: (user: any, error: Error) => void, odataQuery?: string) {
            this.graph.managerForUser(this.userPrincipalName, callback, odataQuery);
        }

        public managerAsync(odataQuery?: string): Promise {
            return this.graph.managerForUserAsync(this.userPrincipalName, odataQuery);
        }      

        public photoValue(callback: (val: any, error: Error) => void) {
            this.graph.photoValueForUser(this.userPrincipalName, callback);
        }

        public photoValueAsync(): Promise {
            return this.graph.photoValueForUserAsync(this.userPrincipalName);
        }

        public calendar(callback: (calendarItems: CalendarEvent[], error: Error) => void, odataQuery?: string) {
            this.graph.calendarForUser(this.userPrincipalName, callback, odataQuery);
        }

        public calendarAsync(odataQuery?: string) {
            return this.graph.calendarForUserAsync(this.userPrincipalName, odataQuery);
        }

    }


    export interface Message {
    }

    export class CalendarEvent {
    }

    export class Contact {
    }

    export interface Group {
    }

    export class Graph {
        private req: XMLHttpRequest = null;
        private state: string = null;
        private nonce: string = null;
        private accessToken: string = null;
        private KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/v1.0/";

        constructor(identityInfo: { identity: Identity });
        constructor(identityInfo: { defaultAccessToken: string });
        constructor(identityInfo: any) {
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }
      
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

        public me(callback: (user: User, error: Error) => void, odataQuery?: string): void {
            var urlString: string = this.buildMeUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUser(urlString, callback);
        }

        public userAsync(userId: string): Promise {
            var d = new Deferred();
            this.user(userId, (user, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(user);
                }
            });
            return d.promise;
        }

        public user(userId: string, callback: (users: any, error: Error) => void): void {
            var urlString: string = this.buildUsersUrl() + "/" + userId;
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

        public users(callback: (users: any, error: Error) => void, odataQuery?: string): void {
            var urlString: string = this.buildUsersUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUsers(urlString, callback);
        }

        //Groups

        public groupAsync(groupId: string): Promise {
            var d = new Deferred();
            this.group(groupId, (group, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(group);
                }
            });
            return d.promise;
        }

        public group(groupId: string, callback: (group: any, error: Error) => void): void {
            var urlString: string = this.buildGroupsUrl() + "/" + groupId;
            this.getGroup(urlString, callback);
        }

        public groups(callback: (groups: any, error: Error) => void, odataQuery?: string): void {
            var urlString: string = this.buildGroupsUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getGroups(urlString, callback);
        }

        public groupsAsync(odataQuery?: string): Promise {
            var d = new Deferred();
            this.groups((groups, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(groups);
                }
            }, odataQuery);
            return d.promise;
        }
        

        // Messages For User
            
        public messagesForUser(userPrincipalName: string, callback: (messages: Message, error: Error) => void, odataQuery?: string): void {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/messages";
            if (odataQuery) urlString += "?" + odataQuery;

            this.getMessages(urlString, (result, error) => {
                callback(result, error);
            }, odataQuery);
        }

        public messagesForUserAsync(userPrincipalName: string, odataQuery?: string): Promise {
            var d = new Deferred();
            this.messagesForUser(userPrincipalName, (messages, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(messages);
                }
            }, odataQuery);
            return d.promise;
        }

        // Calendar For User
            
        public calendarForUser(userPrincipalName: string, callback: (events: CalendarEvent, error: Error) => void, odataQuery?: string): void {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/calendar/events";
            if (odataQuery) urlString += "?" + odataQuery;

            this.getMessages(urlString, (result, error) => {
                callback(result, error);
            }, odataQuery);
        }

        public calendarForUserAsync(userPrincipalName: string, odataQuery?: string): Promise {
            var d = new Deferred();
            this.calendarForUser(userPrincipalName, (events, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(events);
                }
            }, odataQuery);
            return d.promise;
        }

        // Groups/Relationships For User
        public memberOfForUser(userPrincipalName: string, callback: (groups: any, error: Error) => void, odataQuery?: string) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/memberOf";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getGroups(urlString, callback, odataQuery);
        }

        public memberOfForUserAsync(userPrincipalName: string, odataQuery?: string) {
            var d = new Deferred();
            this.memberOfForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public managerForUser(userPrincipalName: string, callback: (manager: any, error: Error) => void, odataQuery?: string) {
            // need odataQuery;
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/manager";
            this.getUser(urlString, callback);
        }

        public managerForUserAsync(userPrincipalName: string, odataQuery?: string) {
            var d = new Deferred();
            this.managerForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public directReportsForUser(userPrincipalName: string, callback: (users: any, error: Error) => void, odataQuery?: string) {
            // Need odata query
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/directReports";
            this.getUsers(urlString, callback);
        }

        public directReportsForUserAsync(userPrincipalName: string, odataQuery?: string) {
            var d = new Deferred();
            this.directReportsForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public photoForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo";
            this.getPhoto(urlString, callback);
        }

        public photoForUserAsync(userPrincipalName: string) {
            var d = new Deferred();
            this.photoForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }

        public photoValueForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo/$value";
            this.getPhotoValue(urlString, callback);
        }

        public photoValueForUserAsync(userPrincipalName: string) {
            var d = new Deferred();
            this.photoValueForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }
    
        //http verbs
        public getAsync(url: string): Promise {
            var d = new Deferred();
            this.get(url, (response, error) => {
                if (!error)
                    d.resolve(response);
                else
                    d.reject(error);
            });
            return d.promise;
        }

        public get(url: string, callback: (response: string, error: Error) => void, responseType?: string): void {
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = (() => {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    if (!responseType)
                        callback(xhr.responseText, null);
                    else
                        callback(xhr.response, null);
                } else if (xhr.readyState === 4 && xhr.status !== 200) {
                    callback(null, this.generateError(xhr));
                }
            });

            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(null, addTokenError);
                }
            });
        }

        private generateError(xhr: XMLHttpRequest): Error {
            var response = new Error();
            response.status = xhr.status;
            response.statusText = xhr.statusText;
            if (xhr.responseType === '' || xhr.responseType === 'text')
                response.text = xhr.responseText;
            else
                response.other = xhr.response;
            return response;

        }

        //Private methods

        private getUsers(urlString, callback: (users: any, error: Error) => void): void {
            this.get(urlString, ((result: string, errorGet: Error) => {

                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }

                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = usersODATA.error;
                    callback(null, errorODATA);
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
                    (<any>users).nextLink = ((callback?: (result: string, error: Error) => void) => {
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

        private getUser(urlString, callback: (user: User, error: Error) => void): void {
            this.get(urlString, (result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var userODATA = JSON.parse(result) ;
                if (userODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = userODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                var user = new User(this, userODATA);
                callback(user, null);
            });

        }

        private addAccessTokenAndSend(xhr: XMLHttpRequest, callback: (error: Error) => void): void {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            } else {
                //Using the integrated Identity object
                this.KurveIdentity.getAccessToken(this.defaultResourceID, ((token: string, error: Error) => {
                    //cache the token
                    
                    if (error)
                        callback(error);
                    else {
                        xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                        xhr.send();
                        callback(null);
                    }
                }));
            }
        }

        // no longer used
        private decorateUserObject(user: any): void {

            user.messages = ((callback: (messages: any, error: Error) => void, odataQuery?: string) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/messages";
                if (odataQuery) urlString += "?" + odataQuery;

                this.getMessages(urlString, (result, error) => {
                    callback(result, error);
                }, odataQuery);
            });


            user.messagesAsync = ((odataQuery?: string) => {
                var d = new Deferred();
                user.messages((result, error) => {
                    if (error) {
                        d.reject(error);
                    } else {
                        d.resolve(result);
                    }
                }, odataQuery);
                return d.promise;
            });

            user.memberOf = ((callback: (groups: any, error: Error) => void, odataQuery?: string) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/memberOf";
                if (odataQuery) urlString += "?" + odataQuery;
                this.getGroups(urlString, callback, odataQuery);
            });

            user.memberOfAsync = ((odataQuery?: string) => {
                var d = new Deferred();
                user.memberOf((result, error) => {
                    if (error) {
                        d.reject(error);
                    } else {
                        d.resolve(result);
                    }
                }, odataQuery);
                return d.promise;
            });

            user.manager = ((callback: (manager: any, error: Error) => void) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/manager";
                this.getUser(urlString, callback);
            });

            user.managerAsync = ((odataQuery?: string) => {
                var d = new Deferred();
                user.manager((result, error) => {
                    if (error) {
                        d.reject(error);
                    } else {
                        d.resolve(result);
                    }
                }, odataQuery);
                return d.promise;
            });

            user.directReports = ((callback: (users: any, error: Error) => void) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/directReports";
                this.getUsers(urlString, callback);
            });

            user.directReportsAsync = ((odataQuery?: string) => {
                var d = new Deferred();
                user.directReports((result, error) => {
                    if (error) {
                        d.reject(error);
                    } else {
                        d.resolve(result);
                    }
                }, odataQuery);
                return d.promise;
            });

            user.photo = ((callback: (photo: any, error: Error) => void) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/photo";
                this.getPhoto(urlString, callback);
            });
            user.photoAsync = (() => {
                var d = new Deferred();
                user.photo((result, error) => {
                    if (error) {
                        d.reject(error);
                    } else {
                        d.resolve(result);
                    }
                });
                return d.promise;
            });
            user.photoValue = ((callback: (photo: any, error: Error) => void) => {
                var urlString = this.buildUsersUrl() + "/" + user.userPrincipalName + "/photo/$value";
                this.getPhotoValue(urlString, callback);
            });
            user.photoValueAsync = (() => {
                var d = new Deferred();
                user.photoValue((result, error) => {
                    if (error) {
                        d.reject(error);
                    } else {
                        d.resolve(result);
                    }
                });
                return d.promise;
            });

        }

        private decorateMessageObject(message: any): void {
        }

        private decorateGroupObject(message: any): void {
        }

        private decoratePhotoObject(message: any): void {
        }

        private getMessages(urlString: string, callback: (messages: any, error: Error) => void, odataQuery?: string): void {

            var url = urlString;
            if (odataQuery) urlString += "?" + odataQuery;
            this.get(url, ((result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }

                var messagesODATA = JSON.parse(result);
                if (messagesODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = messagesODATA.error;
                    callback(null, errorODATA);
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
                    (<any>messages).nextLink = ((callback?: (result: string, error: Error) => void) => {
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

        private getGroups(urlString: string, callback: (groups: any, error: Error) => void, odataQuery?: string): void {

            var url = urlString;
            if (odataQuery) urlString += "?" + odataQuery;
            this.get(url, ((result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var groupsODATA = JSON.parse(result);
                if (groupsODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = groupsODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                var resultsArray = !groupsODATA.value ? [groupsODATA] : groupsODATA.value;

                for (var i: number = 0; i < resultsArray.length; i++) {
                    this.decorateGroupObject(resultsArray[i]);
                }

                var groups = {
                    resultsPage: resultsArray
                };
                var nextLink = groupsODATA['@odata.nextLink'];

                //implement nextLink
                if (nextLink) {
                    (<any>groups).nextLink = ((callback?: (result: string, error: Error) => void) => {
                        var d = new Deferred();
                        this.getGroups(nextLink, ((result, error) => {
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

                callback(groups, null);
            }));
        }

        private getGroup(urlString, callback: (group: any, error: Error) => void): void {
            this.get(urlString, (result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }

                this.decorateGroupObject(ODATA);

                callback(ODATA, null);
            });

        }

        private getPhoto(urlString, callback: (photo: any, error: Error) => void): void {
            this.get(urlString, (result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = ODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                this.decoratePhotoObject(ODATA);

                callback(ODATA, null);
            });
        }

        private getPhotoValue(urlString, callback: (photo: any, error: Error) => void): void {
            this.get(urlString, (result: any, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                callback(result, null);
            }, "blob");
        }
        private buildMeUrl(): string {
            return this.baseUrl + "me";
        }
        private buildUsersUrl(): string {
            return this.baseUrl + "/users";
        }
        private buildGroupsUrl(): string {
            return this.baseUrl + "/groups";
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
