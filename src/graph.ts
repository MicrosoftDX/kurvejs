// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

import { Deferred, Promise, PromiseCallback } from "./promises";
import { Identity, OAuthVersion, Error } from "./identity";
import { UserDataModel, ProfilePhotoDataModel, MessageDataModel, EventDataModel, GroupDataModel, MailFolderDataModel, AttachmentDataModel } from "./models"
import { Collection, User, Users, pathWithQuery } from "./requestbuilder";
/*
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
*/
/*
    export class DataModelWrapper<T> {
        constructor(protected graph: Graph, protected _data: T) {
        }
        get data() { return this._data; }
    }

    export class DataModelListWrapper<T,S> extends DataModelWrapper<T[]>{
        public nextLink: NextLink<S>;
    }

    export class ProfilePhoto extends DataModelWrapper<ProfilePhotoDataModel> {
    }

    export enum EventsEndpoint { events, calendarView }

    export class User extends DataModelWrapper<UserDataModel> {
        // These are all passthroughs to the graph

        public events(callback: PromiseCallback<Events>, odataQuery?: string) {
            this.graph.eventsForUser(this._data.userPrincipalName, EventsEndpoint.events, callback, odataQuery);
        }

        public eventsAsync(odataQuery?: string): Promise<Events, Error> {
            return this.graph.eventsForUserAsync(this._data.userPrincipalName, EventsEndpoint.events, odataQuery);
        }

        public memberOf(callback: PromiseCallback<Groups>, Error, odataQuery?: string) {
            this.graph.memberOfForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public memberOfAsync(odataQuery?: string): Promise<Groups, Error> {
            return this.graph.memberOfForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public messages(callback: PromiseCallback<Messages>, odataQuery?: string) {
            this.graph.messagesForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public messagesAsync(odataQuery?: string): Promise<Messages, Error> {
            return this.graph.messagesForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public manager(callback: PromiseCallback<User>, odataQuery?: string) {
            this.graph.managerForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public managerAsync(odataQuery?: string): Promise<User, Error> {
            return this.graph.managerForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public profilePhoto(callback: PromiseCallback<ProfilePhoto>, odataQuery?: string) {
            this.graph.profilePhotoForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public profilePhotoAsync(odataQuery?: string): Promise<ProfilePhoto, Error> {
            return this.graph.profilePhotoForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public profilePhotoValue(callback: PromiseCallback<any>, odataQuery?: string) {
            this.graph.profilePhotoValueForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public profilePhotoValueAsync(odataQuery?: string): Promise<any, Error> {
            return this.graph.profilePhotoValueForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public calendarView(callback: PromiseCallback<Events>, odataQuery?: string) {
            this.graph.eventsForUser(this._data.userPrincipalName, EventsEndpoint.calendarView, callback, odataQuery);
        }

        public calendarViewAsync(odataQuery?: string): Promise<Events, Error> {
            return this.graph.eventsForUserAsync(this._data.userPrincipalName, EventsEndpoint.calendarView, odataQuery);
        }

        public mailFolders(callback: PromiseCallback<MailFolders>, odataQuery?: string) {
            this.graph.mailFoldersForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public mailFoldersAsync(odataQuery?: string): Promise<MailFolders, Error> {
            return this.graph.mailFoldersForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public message(messageId: string, callback: PromiseCallback<Message>, odataQuery?: string) {
            this.graph.messageForUser(this._data.userPrincipalName, messageId, callback, odataQuery);
        }

        public messageAsync(messageId: string, odataQuery?: string): Promise<Message, Error> {
            return this.graph.messageForUserAsync(this._data.userPrincipalName, messageId, odataQuery);
        }

        public event(eventId: string, callback: PromiseCallback<Event>, odataQuery?: string) {
            this.graph.eventForUser(this._data.userPrincipalName, eventId, callback, odataQuery);
        }

        public eventAsync(eventId: string, odataQuery?: string): Promise<Event, Error> {
            return this.graph.eventForUserAsync(this._data.userPrincipalName, eventId, odataQuery);
        }


        public messageAttachment(messageId: string, attachmentId: string, callback: PromiseCallback<Attachment>, odataQuery?: string) {
            this.graph.messageAttachmentForUser(this._data.userPrincipalName, messageId, attachmentId, callback, odataQuery);
        }

        public messageAttachmentAsync(messageId: string, attachmentId: string, odataQuery?: string): Promise<Attachment, Error> {
            return this.graph.messageAttachmentForUserAsync(this._data.userPrincipalName, messageId, attachmentId, odataQuery);
        }


    }

    export interface NextLink<T> {
        (callback? : PromiseCallback<T>): Promise<T, Error>;
    }

    export class Users extends DataModelListWrapper<User, Users>{
    }

    export class Message extends DataModelWrapper<MessageDataModel>{
    }

    export class Messages extends DataModelListWrapper<Message, Messages>{
    }

    export class Event extends DataModelWrapper<EventDataModel>{
    }

    export class Events extends DataModelListWrapper<Event, Events>{
        constructor(protected graph: Graph, protected endpoint: EventsEndpoint, protected _data: Event[]) {
            super(graph, _data);
        }
    }

    export class Contact {
    }

    export class Group extends DataModelWrapper<GroupDataModel>{
    }

    export class Groups extends DataModelListWrapper<Group, Groups>{
    }

    export class MailFolder extends DataModelWrapper<MailFolderDataModel>{
    }

    export class MailFolders extends DataModelListWrapper<MailFolder, MailFolders>{
    }

	export enum AttachmentType {
		fileAttachment,
		itemAttachment,
		referenceAttachment
	}

    export class Attachment extends DataModelWrapper<AttachmentDataModel>{
        public getType() : AttachmentType {
            switch (this._data['@odata.type']) {
                case "#microsoft.graph.fileAttachment":
                    return AttachmentType.fileAttachment;
                case "#microsoft.graph.itemAttachment":
                    return AttachmentType.itemAttachment;
                case "#microsoft.graph.referenceAttachment":
                    return AttachmentType.referenceAttachment;
            }
        }
    }

    export class Attachments extends DataModelListWrapper<Attachment, Attachments>{
    }
*/
    export class Graph {
        private req: XMLHttpRequest = null;
        private accessToken: string = null;
        private KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/v1.0";

        constructor(identityInfo: { identity: Identity });
        constructor(identityInfo: { defaultAccessToken: string });
        constructor(identityInfo: any) {
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }

        GET = <Model>(path:string, queryT?:string, scopes?:string[]) => (query?:string) => this.Get<Model>(pathWithQuery(path, queryT, query), scopes);
        GETCOLLECTION = <Model>(path:string, queryT?:string, scopes?:string[]) => (query?:string) => this.GetCollection<Model>(pathWithQuery(path, queryT, query), scopes);

        me = new User(this, this.baseUrl);
        user = (userId:string) => new User(this, this.baseUrl, userId);
        users = new Users(this, this.baseUrl);

        public Get<Model>(path:string, scopes?:string[]): Promise<Model, Error> {
            console.log("GET", path);
            var d = new Deferred<Model, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(jsonResult);
            });

            return d.promise;
         }

        public GetCollection<Model>(path:string, scopes?:string[]): Promise<Collection<Model>, Error> {
            console.log("GETCOLLECTION", path);
            var d = new Deferred<Collection<Model>, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                var resultsArray = (jsonResult.value ? jsonResult.value : [jsonResult]) as any[];

                d.resolve({objects:resultsArray});
            });

            return d.promise;
         }
 
        //Only adds scopes when linked to a v2 Oauth of kurve identity
        private scopesForV2(scopes: string[]): string[] {
            if (!this.KurveIdentity)
                return null;
            if (this.KurveIdentity.getCurrentOauthVersion() === OAuthVersion.v1)
                return null;
            else return scopes;
        }
/*
        //Users
        public meAsync(odataQuery?: string): Promise<User, Error> {
            var d = new Deferred<User,Error>();
            this.me((error, user) => error ? d.reject(error) : d.resolve(user), odataQuery);
            return d.promise;
        }

        public me(callback: PromiseCallback<User>, odataQuery?: string): void {
            var scopes = [Scopes.User.Read];
            var urlString = this.buildMeUrl("", odataQuery);
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        }

        public userAsync(userId: string, odataQuery?: string, basicProfileOnly = true): Promise<User, Error> {
            var d = new Deferred<User,Error>();
            this.user(userId, (error, user) => error ? d.reject(error) : d.resolve(user), odataQuery, basicProfileOnly);
            return d.promise;
        }

        public user(userId: string, callback: PromiseCallback<User>, odataQuery?: string, basicProfileOnly = true): void {
            var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
            var urlString = this.buildUsersUrl(userId, odataQuery);
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        }

        public usersAsync(odataQuery?: string, basicProfileOnly = true): Promise<Users, Error> {
            var d = new Deferred<Users,Error>();
            this.users((error, users) => error ? d.reject(error) : d.resolve(users), odataQuery, basicProfileOnly);
            return d.promise;
        }

        public users(callback: PromiseCallback<Users>, odataQuery?: string, basicProfileOnly = true): void {
            var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
            var urlString = this.buildUsersUrl("", odataQuery);
            this.getUsers(urlString, callback, this.scopesForV2(scopes), basicProfileOnly);
        }

        //Groups
        public groupAsync(groupId: string, odataQuery?: string): Promise<Group,Error> {
            var d = new Deferred<Group,Error>();
            this.group(groupId, (error, group) => error ? d.reject(error) : d.resolve(group), odataQuery);
            return d.promise;
        }

        public group(groupId: string, callback: PromiseCallback<Group>, odataQuery?: string): void {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildGroupsUrl(groupId, odataQuery);
            this.getGroup(urlString, callback, this.scopesForV2(scopes));
        }

        public groupsAsync(odataQuery?: string): Promise<Groups, Error> {
            var d = new Deferred<Groups, Error>();
            this.groups((error, groups) => error ? d.reject(error) : d.resolve(groups), odataQuery);
            return d.promise;
        }

        public groups(callback: PromiseCallback<Groups>, odataQuery?: string): void {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildGroupsUrl("", odataQuery);
            this.getGroups(urlString, callback, this.scopesForV2(scopes));
        }

        // Messages For User
        public messageForUserAsync(userPrincipalName: string, messageId: string, odataQuery?: string): Promise<Message, Error> {
            var d = new Deferred<Message, Error>();
            this.messageForUser(userPrincipalName, messageId, (error, message) => error ? d.reject(error) : d.resolve(message), odataQuery);
            return d.promise;
        }

        public messageForUser(userPrincipalName: string, messageId: string, callback: PromiseCallback<Message>, odataQuery?: string): void {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages/" + messageId, odataQuery);
            this.getMessage(urlString, messageId, (error, result) => callback(error, result), this.scopesForV2(scopes));
        }

        public messagesForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error> {
            var d = new Deferred<Messages, Error>();
            this.messagesForUser(userPrincipalName, (error, messages) => error ? d.reject(error) : d.resolve(messages), odataQuery);
            return d.promise;
        }

        public messagesForUser(userPrincipalName: string, callback: PromiseCallback<Messages>, odataQuery?: string): void {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages", odataQuery);
            this.getMessages(urlString, (error, result) => callback(error, result), this.scopesForV2(scopes));
        }

        // MailFolders For User
        public mailFoldersForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<MailFolders, Error> {
            var d = new Deferred<MailFolders, Error>();
            this.mailFoldersForUser(userPrincipalName, (error, messages) => error ? d.reject(error) : d.resolve(messages), odataQuery);
            return d.promise;
        }

        public mailFoldersForUser(userPrincipalName: string, callback: PromiseCallback<MailFolders>, odataQuery?: string): void {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/mailFolders", odataQuery);
            this.getMailFolders(urlString, (error, result) => callback(error, result), this.scopesForV2(scopes));
        }

        // Events For User

        public eventForUserAsync(userPrincipalName: string, eventId: string, odataQuery?: string): Promise<Event, Error> {
            var d = new Deferred<Event, Error>();
            this.eventForUser(userPrincipalName, eventId, (error, event) => error ? d.reject(error) : d.resolve(event), odataQuery);
            return d.promise;
        }

        public eventForUser(userPrincipalName: string, eventId: string, callback: PromiseCallback<Event>, odataQuery?: string): void {
            var scopes = [Scopes.Calendars.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/events/" + eventId, odataQuery);
            this.getEvent(urlString, eventId, (error, result) => callback(error, result), this.scopesForV2(scopes));
        }

        public eventsForUserAsync(userPrincipalName: string, endpoint: EventsEndpoint, odataQuery?: string): Promise<Events, Error> {
            var d = new Deferred<Events, Error>();
            this.eventsForUser(userPrincipalName, endpoint, (error, events) => error ? d.reject(error) : d.resolve(events), odataQuery);
            return d.promise;
        }

        public eventsForUser(userPrincipalName: string, endpoint: EventsEndpoint, callback: (error: Error, messages: Events) => void, odataQuery?: string): void {
            var scopes = [Scopes.Calendars.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/" + EventsEndpoint[endpoint], odataQuery);
            this.getEvents(urlString, endpoint, (error, result) => callback(error, result), this.scopesForV2(scopes));
        }

        // Groups/Relationships For User
        public memberOfForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Groups, Error> {
            var d = new Deferred<Groups, Error>();
            this.memberOfForUser(userPrincipalName, (error, groups) => error ? d.reject(error) : d.resolve(groups), odataQuery);
            return d.promise;
        }

        public memberOfForUser(userPrincipalName: string, callback: PromiseCallback<Groups>, odataQuery?: string) {
            var scopes = [Scopes.Group.ReadAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/memberOf", odataQuery);
            this.getGroups(urlString, callback, this.scopesForV2(scopes));
        }

        public managerForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<User, Error> {
            var d = new Deferred<User, Error>();
            this.managerForUser(userPrincipalName, (error, user) => error ? d.reject(error) : d.resolve(user), odataQuery);
            return d.promise;
        }

        public managerForUser(userPrincipalName: string, callback: PromiseCallback<User>, odataQuery?: string) {
            var scopes = [Scopes.Directory.ReadAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/manager", odataQuery);
            this.getUser(urlString, callback, this.scopesForV2(scopes));
        }

        public directReportsForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Users, Error> {
            var d = new Deferred<Users, Error>();
            this.directReportsForUser(userPrincipalName, (error, users) => error ? d.reject(error) : d.resolve(users), odataQuery);
            return d.promise;
        }

        public directReportsForUser(userPrincipalName: string, callback: PromiseCallback<Users>, odataQuery?: string) {
            var scopes = [Scopes.Directory.ReadAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/directReports", odataQuery);
            this.getUsers(urlString, callback, this.scopesForV2(scopes));
        }

        public profilePhotoForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<ProfilePhoto, Error> {
            var d = new Deferred<ProfilePhoto, Error>();
            this.profilePhotoForUser(userPrincipalName, (error, profilePhoto) => error ? d.reject(error) : d.resolve(profilePhoto), odataQuery);
            return d.promise;
        }

        public profilePhotoForUser(userPrincipalName: string, callback: PromiseCallback<ProfilePhoto>, odataQuery?: string) {
            var scopes = [Scopes.User.ReadBasicAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/photo", odataQuery);
            this.getPhoto(urlString, callback, this.scopesForV2(scopes));
        }

        public profilePhotoValueForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<any, Error> {
            var d = new Deferred<any, Error>();
            this.profilePhotoValueForUser(userPrincipalName, (error, result) => error ? d.reject(error) : d.resolve(result), odataQuery);
            return d.promise;
        }

        public profilePhotoValueForUser(userPrincipalName: string, callback: PromiseCallback<any>, odataQuery?: string) {
            var scopes = [Scopes.User.ReadBasicAll];
            var urlString = this.buildUsersUrl(userPrincipalName + "/photo/$value", odataQuery);
            this.getPhotoValue(urlString, callback, this.scopesForV2(scopes));
        }

        // Message Attachments
        public messageAttachmentsForUserAsync(userPrincipalName: string, messageId: string, odataQuery?: string): Promise<Attachments, Error> {
            var d = new Deferred<any, Error>();
            this.messageAttachmentsForUser(userPrincipalName, messageId, (error, result) => error ? d.reject(error) : d.resolve(result), odataQuery);
            return d.promise;
        }

        public messageAttachmentsForUser(userPrincipalName: string, messageId: string, callback: PromiseCallback<Attachments>, odataQuery?: string): void {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages/" + messageId + "/attachments", odataQuery);
            this.getMessageAttachments(urlString, callback, this.scopesForV2(scopes));
        }

        public messageAttachmentForUserAsync(userPrincipalName: string, messageId: string, attachmentId: string, odataQuery?: string): Promise<Attachment, Error> {
            var d = new Deferred<Attachment,Error>();
            this.messageAttachmentForUser(userPrincipalName, messageId, attachmentId, (error, attachment) => error ? d.reject(error) : d.resolve(attachment), odataQuery);
            return d.promise;
        }

        public messageAttachmentForUser(userPrincipalName: string, messageId: string, attachmentId: string, callback: PromiseCallback<Attachment>, odataQuery?: string): void {
            var scopes = [Scopes.Mail.Read];
            var urlString = this.buildUsersUrl(userPrincipalName + "/messages/" + messageId + "/attachments/" + attachmentId, odataQuery);
            this.getMessageAttachment(urlString, callback, this.scopesForV2(scopes));
        }
*/
        //http verbs
        public getAsync(url: string): Promise<string, Error> {
            var d = new Deferred<string,Error>();
            this.get(url, (error, response) => error ? d.reject(error) : d.resolve(response))
            return d.promise;
        }

        public get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?:string[]): void {
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4)
                    if (xhr.status === 200)
                        callback(null, responseType ? xhr.response : xhr.responseText);
                    else
                        callback(this.generateError(xhr));
            }

            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(addTokenError);
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
/*
        //Private methods

        private getUsers(urlString, callback: PromiseCallback<Users>, scopes?: string[], basicProfileOnly = true): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = usersODATA.error;
                    callback(errorODATA);
                    return;
                }

                var resultsArray = (usersODATA.value ? usersODATA.value : [usersODATA]) as any[];
                var users = new Users(this, resultsArray.map(o => new User(this, o)));
                var nextLink = usersODATA['@odata.nextLink'];
                if (nextLink) {
                    users.nextLink = (callback?: PromiseCallback<Users>) => {
                        var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
                        var d = new Deferred<Users,Error>();
                        this.getUsers(nextLink, (error: Error, result: Users) => {
                            if (callback)
                                callback(error, result);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, this.scopesForV2(scopes), basicProfileOnly);
                        return d.promise;
                    }
                }

                callback(null, users);
            },null,scopes);
        }

        private getUser(urlString, callback: PromiseCallback<User>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }
                var userODATA = JSON.parse(result) ;
                if (userODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = userODATA.error;
                    callback(errorODATA);
                    return;
                }

                var user = new User(this, userODATA);
                callback(null, user);
            },null,scopes);

        }
*/
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
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, ((error: Error, token: string) => {
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
/*
        private getMessage(urlString: string, messageId: string, callback: PromiseCallback<Message>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Error();
                    ODATAError.other = ODATA.error;
                    callback(ODATAError);
                    return;
                }
                var message = new Message(this, ODATA);

                callback(null, message);
            }, null, scopes);

        }

        private getMessages(urlString: string, callback: PromiseCallback<Messages>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var messagesODATA = JSON.parse(result);
                if (messagesODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = messagesODATA.error;
                    callback(errorODATA);
                    return;
                }

                var resultsArray = (messagesODATA.value ? messagesODATA.value : [messagesODATA]) as any[];
                var messages = new Messages(this, resultsArray.map(o => new Message(this, o)));
                var nextLink = messagesODATA['@odata.nextLink'];
                if (nextLink) {
                    messages.nextLink = (callback?: PromiseCallback<Messages>) => {
                        var scopes = [Scopes.Mail.Read];
                        var d = new Deferred<Messages,Error>();
                        this.getMessages(nextLink, (error: Error, messages: Messages) => {
                            if (callback)
                                callback(error, messages);
                            else
                                error ? d.reject(error) : d.resolve(messages);
                        }, this.scopesForV2(scopes));
                        return d.promise;
                    }
                }
                callback(null, messages);
            },null,scopes);
        }

        private getEvent(urlString: string, EventId: string, callback: PromiseCallback<Event>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Error();
                    ODATAError.other = ODATA.error;
                    callback(ODATAError);
                    return;
                }
                var event = new Event(this, ODATA);

                callback(null, event);
            }, null, scopes);

        }

        private getEvents(urlString: string, endpoint: EventsEndpoint, callback: PromiseCallback<Events>, scopes?: string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var odata = JSON.parse(result);
                if (odata.error) {
                    var errorODATA = new Error();
                    errorODATA.other = odata.error;
                    callback(errorODATA);
                    return;
                }

                var resultsArray = (odata.value ? odata.value : [odata]) as any[];
                var events = new Events(this, endpoint, resultsArray.map(o => new Event(this, o)));
                var nextLink = odata['@odata.nextLink'];
                if (nextLink) {
                    events.nextLink = (callback?: PromiseCallback<Events>) => {
                        var scopes = [Scopes.Mail.Read];
                        var d = new Deferred<Events, Error>();
                        this.getEvents(nextLink, endpoint, (error: Error, result: Events) => {
                            if (callback)
                                callback(error, result);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, this.scopesForV2(scopes));
                        return d.promise;
                    }
                }
                callback(null, events);
            }, null, scopes);
        }


        private getGroups(urlString: string, callback: PromiseCallback<Groups>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var groupsODATA = JSON.parse(result);
                if (groupsODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = groupsODATA.error;
                    callback(errorODATA);
                    return;
                }

                var resultsArray = (groupsODATA.value ? groupsODATA.value : [groupsODATA]) as any[];
                var groups = new Groups(this, resultsArray.map(o => new Group(this, o)));
                var nextLink = groupsODATA['@odata.nextLink'];
                if (nextLink) {
                    groups.nextLink = (callback: PromiseCallback<Groups>) => {
                        var scopes = [Scopes.Group.ReadAll];
                        var d = new Deferred<Groups,Error>();
                        this.getGroups(nextLink, (error: Error, result: Groups) => {
                            if (callback)
                                callback(error, result);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, this.scopesForV2(scopes));
                        return d.promise;
                    }
                }

                callback(null, groups);
            },null,scopes);
        }

        private getGroup(urlString: string, callback: PromiseCallback<Group>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Error();
                    ODATAError.other = ODATA.error;
                    callback(ODATAError);
                    return;
                }
                var group = new Group(this, ODATA);

                callback(null, group);
            },null,scopes);

        }

        private getPhoto(urlString, callback: PromiseCallback<ProfilePhoto>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = ODATA.error;
                    callback(errorODATA);
                    return;
                }
                var photo = new ProfilePhoto(this, ODATA);

                callback(null, photo);
            },null,scopes);
        }

        private getPhotoValue(urlString, callback: PromiseCallback<any>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: any) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                callback(result);
            }, "blob",scopes);
        }

        private getMailFolders(urlString, callback: PromiseCallback<MailFolders>, scopes?: string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var odata = JSON.parse(result);
                if (odata.error) {
                    var errorODATA = new Error();
                    errorODATA.other = odata.error;
                    callback(errorODATA);
                }

                var resultsArray:MailFolderDataModel[] = (odata.value ? odata.value : [odata]);
                var mailFolders = new MailFolders(this, resultsArray.map(o => new MailFolder(this, o)));
                var nextLink = odata['@odata.nextLink'];
                if (nextLink) {
                    mailFolders.nextLink = (callback?: PromiseCallback<MailFolders>) => {
                        var scopes = [Scopes.User.ReadAll];
                        var d = new Deferred<MailFolders,Error>();
                        this.getMailFolders(nextLink, (error: Error, result: MailFolders) => {
                            if (callback)
                                callback(error, result);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, this.scopesForV2(scopes));
                        return d.promise;
                    }
                }
                callback(null, mailFolders);
            },null,scopes);
        }


        private getMessageAttachments(urlString: string, callback: PromiseCallback<Attachments>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var attachmentsODATA = JSON.parse(result);
                if (attachmentsODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = attachmentsODATA.error;
                    callback(errorODATA);
                    return;
                }
                var resultsArray = (attachmentsODATA.value ? attachmentsODATA.value : [attachmentsODATA]) as any[];
                var attachments = new Attachments(this, resultsArray.map(o => new Attachment(this, o)));
                var nextLink = attachmentsODATA['@odata.nextLink'];
                if (nextLink) {
                    attachments.nextLink = (callback?: PromiseCallback<Attachments>) => {
                        var scopes = [Scopes.Mail.Read];
                        var d = new Deferred<Attachments,Error>();
                        this.getMessageAttachments(nextLink, (error: Error, attachments: Attachments) => {
                            if (callback)
                                callback(error, attachments);
                            else
                                error ? d.reject(error) : d.resolve(attachments);
                        }, this.scopesForV2(scopes));
                        return d.promise;
                    }
                }

                callback(null, attachments);
            },null,scopes);
        }

        private getMessageAttachment(urlString, callback: PromiseCallback<Attachment>, scopes?:string[]): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Error();
                    ODATAError.other = ODATA.error;
                    callback(ODATAError);
                    return;
                }
                var attachment = new Attachment(this, ODATA);

                callback(null, attachment);
            },null,scopes);
        }

        private buildUrl(root:string, path: string, odataQuery?: string) {
            return this.baseUrl + root + path + (odataQuery ? "?" + odataQuery : "");
        }
        private buildMeUrl(path: string = "", odataQuery?: string) {
            return this.buildUrl("me/", path, odataQuery);
        }
        private buildUsersUrl(path: string = "", odataQuery?: string) {
            return this.buildUrl("users/", path, odataQuery);
        }
        private buildGroupsUrl(path: string = "", odataQuery?: string) {
            return this.buildUrl("groups/", path, odataQuery);
        }
*/
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


/*

var graph = new Graph(new Identity({}));

graph.me.message("123").actions

*/