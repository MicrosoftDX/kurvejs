// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {

    export class ProfilePhotoDataModel {
        public id: string;
        public height: Number;
        public width: Number;        
    }

    export class ProfilePhoto {
        constructor(protected graph: Kurve.Graph, protected _data: ProfilePhotoDataModel) {
        }

        get data() { return this._data; }
    }

    export class UserDataModel {
         public businessPhones;
         public displayName: string;
         public givenName: string;
         public jobTitle: string;
         public mail: string;
         public mobilePhone: string;
         public officeLocation: string;
         public preferredLanguage: string;
         public surname: string;
         public userPrincipalName: string;
         public id: string;
    }

    export class User  {        
        constructor(protected graph: Kurve.Graph, protected _data: UserDataModel) {
        }

        get data() { return this._data; }

        // These are all passthroughs to the graph

        public memberOf(callback: (groups: Groups, Error) => void, Error, odataQuery?: string) {
            this.graph.memberOfForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public memberOfAsync(odataQuery?: string): TypedPromise<Messages> {
            return this.graph.memberOfForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public messages(callback: (messages: Messages, error: Error) => void, odataQuery?: string) {
            this.graph.messagesForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public messagesAsync(odataQuery?: string): TypedPromise<Messages> {
            return this.graph.messagesForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public manager(callback: (user: Kurve.User, error: Error) => void, odataQuery?: string) {
            this.graph.managerForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public managerAsync(odataQuery?: string): TypedPromise<User> {
            return this.graph.managerForUserAsync(this._data.userPrincipalName, odataQuery);
        }      

        public profilePhoto(callback: (photo: ProfilePhoto, error: Error) => void) {
            this.graph.profilePhotoForUser(this._data.userPrincipalName, callback);
        }

        public profilePhotoAsync(): TypedPromise<ProfilePhoto> {
            return this.graph.profilePhotoForUserAsync(this._data.userPrincipalName);
        }

        public profilePhotoValue(callback: (val: any, error: Error) => void) {
            this.graph.profilePhotoValueForUser(this._data.userPrincipalName, callback);
        }

        public profilePhotoValueAsync(): TypedPromise<any> {
            return this.graph.profilePhotoValueForUserAsync(this._data.userPrincipalName);
        }

        public calendar(callback: (calendarItems: CalendarEvents, error: Error) => void, odataQuery?: string) {
            this.graph.calendarForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public calendarAsync(odataQuery?: string) : TypedPromise<CalendarEvents> {
            return this.graph.calendarForUserAsync(this._data.userPrincipalName, odataQuery);
        }

    }

    export class Users {

        public nextLink: (callback?: (users: Kurve.Users, error: Error) => void, odataQuery?: string) => Promise
        constructor(protected graph: Kurve.Graph, protected _data: User[]) {
        }

        get data(): User[] {
            return this._data;
        }
    }
    export class MessageDataModel {
        bccRecipients: string[]
        body: Object
        bodyPreview: string;
        categories: string[]
        ccRecipients: string[]
        changeKey: string;
        conversationId: string;
        createdDateTime: string;
        from: any;
        graph: any;
        hasAttachments: boolean;
        id: string;
        importance: string;
        isDeliveryReceiptRequested: boolean;
        isDraft: boolean;
        isRead: boolean;
        isReadReceiptRequested: boolean;
        lastModifiedDateTime: string;
        parentFolderId: string;
        receivedDateTime: string;
        replyTo: any[]
        sender: any;
        sentDateTime: string;
        subject: string;
        toRecipients: string[]
        webLink: string;
    }

    export class Message {
        constructor(protected graph: Kurve.Graph, protected _data: MessageDataModel) {
        }
        get data() : MessageDataModel {
            return this._data;
        }
    }

    export class Messages {

        public nextLink: (callback?: (messages: Kurve.Messages, error: Error) => void, odataQuery?: string) => Promise
        constructor(protected graph: Kurve.Graph, protected _data: Message[]) {
        }

        get data(): Message[] {
            return this._data;
        }
    }

    export class CalendarEvent {
    }
    export class CalendarEvents {

        public nextLink: (callback?: (events: Kurve.CalendarEvents, error: Error) => void, odataQuery?: string) => Promise
        constructor(protected graph: Kurve.Graph, protected _data: CalendarEvent[]) {
        }

        get data(): CalendarEvent[] {
            return this._data;
        }
    }

    export class Contact {
    }

    export class GroupDataModel {
        public id: string;
        public description: string;
        public displayName: string;
        public groupTypes: string[];
        public mail: string;
        public mailEnabled: Boolean;
        public mailNickname: string;
        public onPremisesLastSyncDateTime: Date;
        public onPremisesSecurityIdentifier: string;
        public onPremisesSyncEnabled: Boolean;
        public proxyAddresses: string[];
        public securityEnabled: Boolean;
        public visibility: string;      
    }

    export class Group {
        constructor(protected graph: Kurve.Graph, protected _data: GroupDataModel) {
        }

        get data() { return this._data; }

    }

    export class Groups {

        public nextLink: (callback?: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) => TypedPromise<Groups>
        constructor(protected graph: Kurve.Graph, protected _data: Group[]) {
        }

        get data(): Group[] {
            return this._data;
        }
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
        public meAsync(odataQuery?: string): TypedPromise<User> {
            var d = new TypedDeferred<User>();
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

        public userAsync(userId: string): TypedPromise<User> {
            var d = new TypedDeferred<User>();
            this.user(userId, (user, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(user);
                }
            });
            return d.promise;
        }

        public user(userId: string, callback: (user: Kurve.User, error: Error) => void): void {
            var urlString: string = this.buildUsersUrl() + "/" + userId;
            this.getUser(urlString, callback);
        }

        public usersAsync(odataQuery?: string): TypedPromise<Users> {
            var d = new TypedDeferred<Users>();
            this.users((users, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(users);
                }
            }, odataQuery);
            return d.promise;
        }

        public users(callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string): void {
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
            
        public messagesForUser(userPrincipalName: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/messages";
            if (odataQuery) urlString += "?" + odataQuery;

            this.getMessages(urlString, (result, error) => {
                callback(result, error);
            }, odataQuery);
        }

        public messagesForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<Messages> {
            var d = new TypedDeferred<Messages>();
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
        // // To BE IMPLEMENTED
        //    var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/calendar/events";
        //    if (odataQuery) urlString += "?" + odataQuery;

        //    this.getMessages(urlString, (result, error) => {
        //        callback(result, error);
        //    }, odataQuery);
        }

        public calendarForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<CalendarEvents> {
            var d = new TypedDeferred<CalendarEvents>();
            // // To BE IMPLEMENTED
            //    this.calendarForUser(userPrincipalName, (events, error) => {
            //        if (error) {
            //            d.reject(error);
            //        } else {
            //            d.resolve(events);
            //        }
            //    }, odataQuery);
            return d.promise;
        }

        // Groups/Relationships For User
        public memberOfForUser(userPrincipalName: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/memberOf";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getGroups(urlString, callback, odataQuery);
        }

        public memberOfForUserAsync(userPrincipalName: string, odataQuery?: string) : TypedPromise<Messages> {
            var d = new TypedDeferred<Messages>();
            this.memberOfForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public managerForUser(userPrincipalName: string, callback: (manager: Kurve.User, error: Error) => void, odataQuery?: string) {
            // need odataQuery;
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/manager";
            this.getUser(urlString, callback);
        }

        public managerForUserAsync(userPrincipalName: string, odataQuery?: string) : TypedPromise<User> {
            var d = new TypedDeferred<User>();
            this.managerForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public directReportsForUser(userPrincipalName: string, callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string) {
            // Need odata query
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/directReports";
            this.getUsers(urlString, callback);
        }

        public directReportsForUserAsync(userPrincipalName: string, odataQuery?: string) : TypedPromise<Users> {
            var d = new TypedDeferred<Users>();
            this.directReportsForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public profilePhotoForUser(userPrincipalName: string, callback: (photo: ProfilePhoto, error: Error) => void) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo";
            this.getPhoto(urlString, callback);
        }

        public profilePhotoForUserAsync(userPrincipalName: string): TypedPromise<ProfilePhoto> {
            var d = new TypedDeferred < ProfilePhoto>();
            this.profilePhotoForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }

        public profilePhotoValueForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo/$value";
            this.getPhotoValue(urlString, callback);
        }

        public profilePhotoValueForUserAsync(userPrincipalName: string) : TypedPromise<any>{
            var d = new TypedDeferred<any>();
            this.profilePhotoValueForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }
    
        //http verbs
        public getAsync(url: string): TypedPromise<string> {
            var d = new TypedDeferred<string>();
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

        private getUsers(urlString, callback: (users: Kurve.Users, error: Error) => void): void {
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

                var resultsArray = (usersODATA.value ? usersODATA.value : [usersODATA]) as any[];
                var users = new Kurve.Users(this, resultsArray.map(o => {
                    return new User(this, o);
                }));

                //implement nextLink
                var nextLink = usersODATA['@odata.nextLink'];

                if (nextLink) {
                    users.nextLink = ((callback?: (result: Users, error: Error) => void) => {
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

        private getMessages(urlString: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void {

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

                var resultsArray = (messagesODATA.value ? messagesODATA.value : [messagesODATA]) as any[];
                var messages = new Kurve.Messages(this, resultsArray.map(o => {
                    return new Message(this, o);
                }));
                if (messagesODATA['@odata.nextLink']) {
                    messages.nextLink = (callback?: (messages: Kurve.Messages, error: Error) => void, odataQuery?: string) => {
                        var d = new Deferred();
                        this.getMessages(messagesODATA['@odata.nextLink'], (messages, error) => {
                            if (callback)
                                callback(messages, error);
                            else if (error)
                                d.reject(error);
                            else
                                d.resolve(messages);
                        }, odataQuery);
                        return d.promise;

                    };
                }
                callback(messages,  null);
            }));
        }

        private getGroups(urlString: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string): void {

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

                var resultsArray = (groupsODATA.value ? groupsODATA.value : [groupsODATA]) as any[];
                var groups = new Kurve.Groups(this, resultsArray.map(o => {
                    return new Group(this, o);
                }));

                var nextLink = groupsODATA['@odata.nextLink'];

                //implement nextLink
                if (nextLink) {
                    groups.nextLink = ((callback?: (result: Kurve.Groups, error: Error) => void) => {
                        var d = new TypedDeferred<Groups>();
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

        private getGroup(urlString, callback: (group: Kurve.Group, error: Error) => void): void {
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
                var group = new Kurve.Group(this, ODATA);

                callback(group, null);
            });

        }

        private getPhoto(urlString, callback: (photo: ProfilePhoto, error: Error) => void): void {
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
                var photo = new ProfilePhoto(this, ODATA);

                callback(photo, null);
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
