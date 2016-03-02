// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {
    export module Scopes {
        class Util {
            static rootUrl = "https://graph.microsoft.com/";
        }
        export class General {
            public static OpenId: string = "openid";
            public static OfflineAccess: string = "offline_access";
        }
        export class User {
            public static Read: string = Util.rootUrl + "User.Read";
            public static ReadWrite: string = Util.rootUrl + "User.ReadWrite";
            public static ReadBasicAll: string = Util.rootUrl + "User.ReadBasic.All";
            public static ReadAll: string = Util.rootUrl + "User.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "User.ReadWrite.All";
        }
        export class Contacts {
            public static Read: string = Util.rootUrl + "Contacts.Read";
            public static ReadWrite: string = Util.rootUrl + "Contacts.ReadWrite";
        }
        export class Directory {
            public static ReadAll: string = Util.rootUrl + "Directory.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "Directory.ReadWrite.All";
            public static AccessAsUserAll: string = Util.rootUrl + "Directory.AccessAsUser.All";
        }
        export class Group {
            public static ReadAll: string = Util.rootUrl + "Group.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "Group.ReadWrite.All";
            public static AccessAsUserAll: string = Util.rootUrl + "Directory.AccessAsUser.All";
        }
        export class Mail {
            public static Read: string = Util.rootUrl + "Mail.Read";
            public static ReadWrite: string = Util.rootUrl + "Mail.ReadWrite";
            public static Send: string = Util.rootUrl + "Mail.Send";
        }
        export class Calendars {
            public static Read: string = Util.rootUrl + "Calendars.Read";
            public static ReadWrite: string = Util.rootUrl + "Calendars.ReadWrite";
        }
        export class Files {
            public static Read: string = Util.rootUrl + "Files.Read";
            public static ReadAll: string = Util.rootUrl + "Files.Read.All";
            public static ReadWrite: string = Util.rootUrl + "Files.ReadWrite";
            public static ReadWriteAppFolder: string = Util.rootUrl + "Files.ReadWrite.AppFolder";
            public static ReadWriteSelected: string = Util.rootUrl + "Files.ReadWrite.Selected";
        }
        export class Tasks {
            public static ReadWrite: string = Util.rootUrl + "Tasks.ReadWrite";
        }
        export class People {
            public static Read: string = Util.rootUrl + "People.Read";
            public static ReadWrite: string = Util.rootUrl + "People.ReadWrite";
        }
        export class Notes {
            public static Create: string = Util.rootUrl + "Notes.Create";
            public static ReadWriteCreatedByApp: string = Util.rootUrl + "Notes.ReadWrite.CreatedByApp";
            public static Read: string = Util.rootUrl + "Notes.Read";
            public static ReadAll: string = Util.rootUrl + "Notes.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "Notes.ReadWrite.All";
        }


    }

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
        public businessPhones: string;
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

    export class User {
        private graph: Kurve.Graph;
        private _data: Kurve.UserDataModel;

        constructor(graph: Kurve.Graph, _data: UserDataModel) {
            this.graph = graph;
            this._data = _data;
        }

        get data() { return this._data; }

        // These are all passthroughs to the graph

        public events(callback: (items: Events, error: Error) => void, odataQuery?: string) {
            this.graph.eventsForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public eventsAsync(odataQuery?: string): Promise<Events, Error> {
            return this.graph.eventsForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public memberOf(callback: (groups: Groups, Error) => void, Error, odataQuery?: string) {
            this.graph.memberOfForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public memberOfAsync(odataQuery?: string): Promise<Messages, Error> {
            return this.graph.memberOfForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public messages(callback: (messages: Messages, error: Error) => void, odataQuery?: string) {
            this.graph.messagesForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public messagesAsync(odataQuery?: string): Promise<Messages, Error> {
            return this.graph.messagesForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public manager(callback: (user: Kurve.User, error: Error) => void, odataQuery?: string) {
            this.graph.managerForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public managerAsync(odataQuery?: string): Promise<User, Error> {
            return this.graph.managerForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public profilePhoto(callback: (photo: ProfilePhoto, error: Error) => void) {
            this.graph.profilePhotoForUser(this._data.userPrincipalName, callback);
        }

        public profilePhotoAsync(): Promise<ProfilePhoto, Error> {
            return this.graph.profilePhotoForUserAsync(this._data.userPrincipalName);
        }

        public profilePhotoValue(callback: (val: any, error: Error) => void) {
            this.graph.profilePhotoValueForUser(this._data.userPrincipalName, callback);
        }

        public profilePhotoValueAsync(): Promise<any, Error> {
            return this.graph.profilePhotoValueForUserAsync(this._data.userPrincipalName);
        }

        public calendar(callback: (calendarItems: Events, error: Error) => void, odataQuery?: string) {
            this.graph.eventsForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public calendarAsync(odataQuery?: string): Promise<Events, Error> {
            return this.graph.eventsForUserAsync(this._data.userPrincipalName, odataQuery);
        }

    }

    export class Users {

        public nextLink: (callback?: (users: Kurve.Users, error: Error) => void, odataQuery?: string) => Promise<Users, Error>
        constructor(protected graph: Kurve.Graph, protected _data: User[]) {
        }

        get data(): User[] {
            return this._data;
        }
    }

    export interface ItemBody {
        contentType: string;
        content: string;
    }

    export interface EmailAddress {
        name: string;
        address: string;
    }

    export interface Recipient {
        emailAddress: EmailAddress;
    }     

    export class MessageDataModel {
        bccRecipients: Recipient[];
        body: ItemBody;
        bodyPreview: string;
        categories: string[]
        ccRecipients: Recipient[];
        changeKey: string;
        conversationId: string;
        createdDateTime: string;
        from: Recipient;
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
        replyTo: Recipient[]
        sender: Recipient;
        sentDateTime: string;
        subject: string;
        toRecipients: Recipient[];
        webLink: string;
    }

    export class Message {
        constructor(protected graph: Kurve.Graph, protected _data: MessageDataModel) {
        }
        get data(): MessageDataModel {
            return this._data;
        }
    }

    export class Messages {

        public nextLink: (callback?: (messages: Messages, error: Error) => void, odataQuery?: string) => Promise<Messages, Error>
        constructor(protected graph: Kurve.Graph, protected _data: Message[]) {
        }

        get data(): Message[] {
            return this._data;
        }
    }

    export interface Attendee {
        status: ResponseStatus;
        type: string;
        emailAddress: EmailAddress;
    }

    export interface DateTimeTimeZone {
        dateTime: string;
        timeZone: string;
    }

    export interface PatternedRecurrence { }

    export interface ResponseStatus {
        response: string;
        time: string
    }

    export interface Location {
        displayName: string;
        address: any;
    }

    export class EventDataModel {
        attendees: Attendee[];
        body: ItemBody;
        bodyPreview: string;
        categories: string[];
        changeKey: string;
        createdDateTime: string;
        end: DateTimeTimeZone;
        hasAttachments: boolean;
        iCalUId: string;
        id: string;
        IDBCursor: string;
        importance: string;
        isAllDay: boolean;
        isCancelled: boolean;
        isOrganizer: boolean;
        isReminderOn: boolean;
        lastModifiedDateTime: string;
        location: Location;
        organizer: Recipient;
        originalEndTimeZone: string;
        originalStartTimeZone: string;
        recurrence: PatternedRecurrence;
        reminderMinutesBeforeStart: number;
        responseRequested: boolean;
        responseStatus: ResponseStatus;
        sensitivity: string;
        seriesMasterId: string;
        showAs: string;
        start: DateTimeTimeZone;
        subject: string;
        type: string;
        webLink: string;
    }

    export class Event {
        constructor(protected graph: Kurve.Graph, protected _data: EventDataModel) {
        }
        get data(): EventDataModel {
            return this._data;
        }
    }

    export class Events {
        public nextLink: (callback?: (events: Events, error: Error) => void, odataQuery?: string) => Promise<Events, Error>
        constructor(protected graph: Kurve.Graph, protected _data: Event[]) {
        }

        get data(): Event[] {
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

        public nextLink: (callback?: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) => Promise<Groups, Error>
        constructor(protected graph: Kurve.Graph, protected _data: Group[]) {
        }

        get data(): Group[] {
            return this._data;
        }
    }

    export class Graph {
        private req: XMLHttpRequest = null;
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
      
        //Only adds scopes when linked to a v2 Oauth of kurve identity
        private scopesForV2(scopes: string[]): string[] {
            if (!this.KurveIdentity)
                return null;
            if (this.KurveIdentity.getCurrentOauthVersion() === OAuthVersion.v1)
                return null;
            else return scopes;
        }

        //Users
        public meAsync(odataQuery?: string): Promise<User, Error> {
            var d = new Deferred<User,Error>();
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
            var scopes = [Scopes.User.Read];
            var urlString: string = this.buildMeUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            if (odataQuery) urlString += "?" + odataQuery;
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        }

        public userAsync(userId: string, odataQuery?: string, basicProfileOnly = true): Promise<User, Error> {
            var d = new Deferred<User,Error>();
            this.user(userId, (user, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(user);
                }
            }, odataQuery, basicProfileOnly);
            return d.promise;
        }

        public user(userId: string, callback: (user: Kurve.User, error: Error) => void, odataQuery?: string, basicProfileOnly = true): void {
            var scopes = [];
            if (basicProfileOnly)
                scopes = [Scopes.User.ReadBasicAll];
            else
                scopes = [Scopes.User.ReadAll];

            var urlString: string = this.buildUsersUrl() + "/" + userId;
            if (odataQuery) urlString += "?" + odataQuery;
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        }

        public usersAsync(odataQuery?: string, basicProfileOnly = true): Promise<Users, Error> {
            var d = new Deferred<Users,Error>();
            this.users((users, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(users);
                }
            }, odataQuery, basicProfileOnly);
            return d.promise;
        }

        public users(callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string, basicProfileOnly = true): void {
            var scopes = [];
            if (basicProfileOnly)
                scopes = [Scopes.User.ReadBasicAll];
            else
                scopes = [Scopes.User.ReadAll];
            var urlString: string = this.buildUsersUrl() + "/";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getUsers(urlString, callback, this.scopesForV2(scopes), basicProfileOnly);
        }

        //Groups
        public groupAsync(groupId: string, odataQuery?: string): Promise<Group,Error> {
            var d = new Deferred<Group,Error>();
            this.group(groupId, (group, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(group);
                }
            }, odataQuery);
            return d.promise;
        }

        public group(groupId: string, callback: (group: any, error: Error) => void, odataQuery?: string): void {
            var scopes = [Scopes.Group.ReadAll];
            var urlString: string = this.buildGroupsUrl() + "/" + groupId;
            if (odataQuery) urlString += "?" + odataQuery;
            this.getGroup(urlString, callback, this.scopesForV2(scopes));
        }

        public groupsAsync(odataQuery?: string): Promise<Groups, Error> {
            var d = new Deferred<Groups, Error>();

            this.groups((groups, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(groups);
                }
            }, odataQuery);
            return d.promise;
        }

        public groups(callback: (groups: any, error: Error) => void, odataQuery?: string): void {
            var scopes = [Scopes.Group.ReadAll];
            var urlString: string = this.buildGroupsUrl() + "/";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getGroups(urlString, callback, odataQuery, this.scopesForV2(scopes));
        }      

        // Messages For User
        public messagesForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error> {
            var d = new Deferred<Messages, Error>();
            this.messagesForUser(userPrincipalName, (messages, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(messages);
                }
            }, odataQuery);
            return d.promise;
        }

        public messagesForUser(userPrincipalName: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/messages";
            if (odataQuery) urlString += "?" + odataQuery;

            this.getMessages(urlString, (result, error) => {
                callback(result, error);
            }, odataQuery, this.scopesForV2(scopes));
        }


        // Messages For User
        public eventsForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Events, Error> {
            var d = new Deferred<Events, Error>();
            this.eventsForUser(userPrincipalName, (items, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(items);
                }
            }, odataQuery);
            return d.promise;
        }

        public eventsForUser(userPrincipalName: string, callback: (messages: Events, error: Error) => void, odataQuery?: string): void {
            var scopes = [Scopes.Calendars.Read];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/events";
            if (odataQuery) urlString += "?" + odataQuery;

            this.getEvents(urlString, (result, error) => {
                callback(result, error);
            }, odataQuery, this.scopesForV2(scopes));
        }

        // Groups/Relationships For User
        public memberOfForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error> {
            var d: any = new Deferred<Messages, Error>();
            this.memberOfForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public memberOfForUser(userPrincipalName: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/memberOf";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getGroups(urlString, callback, odataQuery, this.scopesForV2(scopes));
        }

        public managerForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<User, Error> {
            var d = new Deferred<User, Error>();
            this.managerForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public managerForUser(userPrincipalName: string, callback: (manager: Kurve.User, error: Error) => void, odataQuery?: string) {
            var scopes = [Scopes.Directory.ReadAll];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/manager";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        }

        public directReportsForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Users, Error> {
            var d = new Deferred<Users, Error>();
            this.directReportsForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public directReportsForUser(userPrincipalName: string, callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string) {
            var scopes = [Scopes.Directory.ReadAll];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/directReports";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getUsers(urlString, callback, this.scopesForV2(scopes));
        }

        public profilePhotoForUserAsync(userPrincipalName: string): Promise<ProfilePhoto, Error> {
            var d = new Deferred<ProfilePhoto, Error>();

            this.profilePhotoForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }

        public profilePhotoForUser(userPrincipalName: string, callback: (photo: ProfilePhoto, error: Error) => void) {
            var scopes = [Scopes.User.ReadBasicAll];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo";
            this.getPhoto(urlString, callback, this.scopesForV2(scopes));
        }

        public profilePhotoValueForUserAsync(userPrincipalName: string): Promise<any, Error> {
            var d = new Deferred<any, Error>();
            this.profilePhotoValueForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }

        public profilePhotoValueForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void) {
            var scopes = [Scopes.User.ReadBasicAll];
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo/$value";
            this.getPhotoValue(urlString, callback, this.scopesForV2(scopes));
        }
    
        //http verbs
        public getAsync(url: string): Promise<string, Error> {
            var d = new Deferred<string,Error>();
            this.get(url, (response, error) => {
                if (!error) {
                    d.resolve(response);
                }
                else {
                    d.reject(error);
                }
            });
            return d.promise;
        }

        public get(url: string, callback: (response: string, error: Error) => void, responseType?: string, scopes?:string[]): void {
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
            }, scopes);
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

        private getUsers(urlString, callback: (users: Kurve.Users, error: Error) => void, scopes?: string[], basicProfileOnly = true): void {

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

                        var scopes = [];
                        if (basicProfileOnly)
                            scopes = [Scopes.User.ReadBasicAll];
                        else
                            scopes = [Scopes.User.ReadAll];

                        var d = new Deferred<Users,Error>();
                        this.getUsers(nextLink, ((result, error) => {
                            if (callback)
                                callback(result, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(result);
                            }
                        }), this.scopesForV2(scopes), basicProfileOnly);
                        return d.promise;
                    });
                }

                callback(users, null);
            }),null,scopes);
        }

        private getUser(urlString, callback: (user: User, error: Error) => void, scopes?:string[]): void {
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
            },null,scopes);

        }

        private addAccessTokenAndSend(xhr: XMLHttpRequest, callback: (error: Error) => void, scopes?:string[]): void {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            } else {
                //Using the integrated Identity object

                if (scopes) {
                    //v2 scope based tokens
                    this.KurveIdentity.getAccessTokenForScopes(scopes,false, ((token: string, error: Error) => {                   
                        if (error)
                            callback(error);
                        else {
                            xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                            xhr.send();
                            callback(null);
                        }
                    }));

                }
                else {
                    //v1 resource based tokens
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, ((token: string, error: Error) => {                  
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
        }

        private getMessages(urlString: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string, scopes?:string[]): void {

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
                    messages.nextLink = (callback?: (messages: Messages, error: Error) => void, odataQuery?: string) => {

                        var scopes = [Scopes.Mail.Read];
                      

                        var d = new Deferred<Messages,Error>();
                        
                        this.getMessages(messagesODATA['@odata.nextLink'], (messages, error) => {
                            if (callback)
                                callback(messages, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(messages);
                            }
                        }, odataQuery, this.scopesForV2(scopes));
                        return d.promise;

                    };
                }
                callback(messages,  null);
            }),null,scopes);
        }

        private getEvents(urlString: string, callback: (events: Events, error: Error) => void, odataQuery?: string, scopes?: string[]): void {

            var url = urlString;
            if (odataQuery) urlString += "?" + odataQuery;
            this.get(url, ((result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }

                var odata = JSON.parse(result);
                if (odata.error) {
                    var errorODATA = new Error();
                    errorODATA.other = odata.error;
                    callback(null, errorODATA);
                    return;
                }

                var resultsArray = (odata.value ? odata.value : [odata]) as any[];
                var items = new Kurve.Events(this, resultsArray.map(o => {
                    return new Event(this, o);
                }));
                if (odata['@odata.nextLink']) {
                    items.nextLink = (callback?: (cbEvents: Events, error: Error) => void, odataQuery?: string) => {

                        var scopes = [Scopes.Mail.Read];

                        var d = new Deferred<Events, Error>();

                        this.getEvents(odata['@odata.nextLink'], (stuff, error) => {
                            if (callback)
                                callback(stuff, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(stuff);
                            }
                        }, odataQuery, this.scopesForV2(scopes));
                        return d.promise;

                    };
                }
                callback(items, null);
            }), null, scopes);
        }


        private getGroups(urlString: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string,scopes?:string[]): void {

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
                    groups.nextLink = ((callback?: (result: Groups, error: Error) => void) => {

                        var scopes = [Scopes.Group.ReadAll];
                        var d = new Deferred<Groups,Error>();
                        this.getGroups(nextLink, ((result, error) => {
                            if (callback)
                                callback(result, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(result);
                            }
                        }), odataQuery, this.scopesForV2(scopes));
                        return d.promise;
                    });
                }

                callback(groups, null);
            }),null,scopes);
        }

        private getGroup(urlString, callback: (group: Kurve.Group, error: Error) => void,scopes?:string[]): void {
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
            },null,scopes);

        }

        private getPhoto(urlString, callback: (photo: ProfilePhoto, error: Error) => void, scopes?:string[]): void {
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
            },null,scopes);
        }

        private getPhotoValue(urlString, callback: (photo: any, error: Error) => void,scopes?:string[]): void {
            this.get(urlString, (result: any, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                callback(result, null);
            }, "blob",scopes);
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
