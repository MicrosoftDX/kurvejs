declare module 'Kurve/src/identity' {
	import { Promise, PromiseCallback } from 'Kurve/src/promises';
	export enum OAuthVersion {
	    v1 = 1,
	    v2 = 2,
	}
	export class Error {
	    status: number;
	    statusText: string;
	    text: string;
	    other: any;
	}
	export interface TokenStorage {
	    add(key: string, token: any): any;
	    remove(key: string): any;
	    getAll(): any[];
	    clear(): any;
	}
	export class IdToken {
	    Token: string;
	    IssuerIdentifier: string;
	    SubjectIdentifier: string;
	    Audience: string;
	    Expiry: Date;
	    UPN: string;
	    TenantId: string;
	    FamilyName: string;
	    GivenName: string;
	    Name: string;
	    PreferredUsername: string;
	    FullToken: any;
	}
	export interface IdentitySettings {
	    clientId: string;
	    tokenProcessingUri: string;
	    version: OAuthVersion;
	    tokenStorage?: TokenStorage;
	}
	export class Identity {
	    clientId: string;
	    private state;
	    private version;
	    private nonce;
	    private idToken;
	    private loginCallback;
	    private getTokenCallback;
	    private tokenProcessorUrl;
	    private tokenCache;
	    private refreshTimer;
	    private policy;
	    constructor(identitySettings: IdentitySettings);
	    checkForIdentityRedirect(): boolean;
	    private decodeIdToken(idToken);
	    private decodeAccessToken(accessToken, resource?, scopes?);
	    getIdToken(): any;
	    isLoggedIn(): boolean;
	    private renewIdToken();
	    getCurrentOauthVersion(): OAuthVersion;
	    getAccessTokenAsync(resource: string): Promise<string, Error>;
	    getAccessToken(resource: string, callback: PromiseCallback<string>): void;
	    getAccessTokenForScopesAsync(scopes: string[], promptForConsent?: boolean): Promise<string, Error>;
	    getAccessTokenForScopes(scopes: string[], promptForConsent: boolean, callback: (token: string, error: Error) => void): void;
	    loginAsync(loginSettings?: {
	        scopes?: string[];
	        policy?: string;
	        tenant?: string;
	    }): Promise<void, Error>;
	    login(callback: (error: Error) => void, loginSettings?: {
	        scopes?: string[];
	        policy?: string;
	        tenant?: string;
	    }): void;
	    loginNoWindowAsync(toUrl?: string): Promise<void, Error>;
	    loginNoWindow(callback: (error: Error) => void, toUrl?: string): void;
	    logOut(): void;
	    private base64Decode(encodedString);
	    private generateNonce();
	}

}
declare module 'Kurve/src/promises' {
	export interface PromiseCallback<T> {
	    (T: any, Error: any): void;
	}
	export class Deferred<T, E> {
	    private _dispatcher;
	    constructor();
	    constructor(dispatcher: (closure: () => void) => void);
	    private DispatchDeferred(closure);
	    then(successCB: any, errorCB: any): any;
	    resolve(value?: T): Deferred<T, E>;
	    resolve(value?: Promise<T, E>): Deferred<T, E>;
	    private _resolve(value);
	    reject(error?: E): Deferred<T, E>;
	    private _reject(error?);
	    promise: Promise<T, E>;
	    private _stack;
	    private _state;
	    private _value;
	    private _error;
	}
	export class Promise<T, E> implements Promise<T, E> {
	    private _deferred;
	    constructor(_deferred: Deferred<T, E>);
	    then<R>(successCallback?: (result: T) => R, errorCallback?: (error: E) => R): Promise<R, E>;
	    fail<R>(errorCallback?: (error: E) => R): Promise<R, E>;
	}

}
declare module 'Kurve/src/graph' {
	import { Promise, PromiseCallback } from 'Kurve/src/promises';
	import { Identity, Error } from 'Kurve/src/identity';
	export module Scopes {
	    class General {
	        static OpenId: string;
	        static OfflineAccess: string;
	    }
	    class User {
	        static Read: string;
	        static ReadWrite: string;
	        static ReadBasicAll: string;
	        static ReadAll: string;
	        static ReadWriteAll: string;
	    }
	    class Contacts {
	        static Read: string;
	        static ReadWrite: string;
	    }
	    class Directory {
	        static ReadAll: string;
	        static ReadWriteAll: string;
	        static AccessAsUserAll: string;
	    }
	    class Group {
	        static ReadAll: string;
	        static ReadWriteAll: string;
	        static AccessAsUserAll: string;
	    }
	    class Mail {
	        static Read: string;
	        static ReadWrite: string;
	        static Send: string;
	    }
	    class Calendars {
	        static Read: string;
	        static ReadWrite: string;
	    }
	    class Files {
	        static Read: string;
	        static ReadAll: string;
	        static ReadWrite: string;
	        static ReadWriteAppFolder: string;
	        static ReadWriteSelected: string;
	    }
	    class Tasks {
	        static ReadWrite: string;
	    }
	    class People {
	        static Read: string;
	        static ReadWrite: string;
	    }
	    class Notes {
	        static Create: string;
	        static ReadWriteCreatedByApp: string;
	        static Read: string;
	        static ReadAll: string;
	        static ReadWriteAll: string;
	    }
	}
	export class DataModelWrapper<T> {
	    protected graph: Graph;
	    protected _data: T;
	    constructor(graph: Graph, _data: T);
	    data: T;
	}
	export class DataModelListWrapper<T, S> extends DataModelWrapper<T[]> {
	    nextLink: NextLink<S>;
	}
	export class ProfilePhotoDataModel {
	    id: string;
	    height: Number;
	    width: Number;
	}
	export class ProfilePhoto extends DataModelWrapper<ProfilePhotoDataModel> {
	}
	export class UserDataModel {
	    businessPhones: string;
	    displayName: string;
	    givenName: string;
	    jobTitle: string;
	    mail: string;
	    mobilePhone: string;
	    officeLocation: string;
	    preferredLanguage: string;
	    surname: string;
	    userPrincipalName: string;
	    id: string;
	}
	export enum EventsEndpoint {
	    events = 0,
	    calendarView = 1,
	}
	export class User extends DataModelWrapper<UserDataModel> {
	    events(callback: PromiseCallback<Events>, odataQuery?: string): void;
	    eventsAsync(odataQuery?: string): Promise<Events, Error>;
	    memberOf(callback: PromiseCallback<Groups>, Error: any, odataQuery?: string): void;
	    memberOfAsync(odataQuery?: string): Promise<Groups, Error>;
	    messages(callback: PromiseCallback<Messages>, odataQuery?: string): void;
	    messagesAsync(odataQuery?: string): Promise<Messages, Error>;
	    manager(callback: PromiseCallback<User>, odataQuery?: string): void;
	    managerAsync(odataQuery?: string): Promise<User, Error>;
	    profilePhoto(callback: PromiseCallback<ProfilePhoto>, odataQuery?: string): void;
	    profilePhotoAsync(odataQuery?: string): Promise<ProfilePhoto, Error>;
	    profilePhotoValue(callback: PromiseCallback<any>, odataQuery?: string): void;
	    profilePhotoValueAsync(odataQuery?: string): Promise<any, Error>;
	    calendarView(callback: PromiseCallback<Events>, odataQuery?: string): void;
	    calendarViewAsync(odataQuery?: string): Promise<Events, Error>;
	    mailFolders(callback: PromiseCallback<MailFolders>, odataQuery?: string): void;
	    mailFoldersAsync(odataQuery?: string): Promise<MailFolders, Error>;
	    message(messageId: string, callback: PromiseCallback<Message>, odataQuery?: string): void;
	    messageAsync(messageId: string, odataQuery?: string): Promise<Message, Error>;
	    event(eventId: string, callback: PromiseCallback<Message>, odataQuery?: string): void;
	    eventAsync(eventId: string, odataQuery?: string): Promise<Event, Error>;
	    messageAttachment(messageId: string, attachmentId: string, callback: PromiseCallback<Attachment>, odataQuery?: string): void;
	    messageAttachmentAsync(messageId: string, attachmentId: string, odataQuery?: string): Promise<Attachment, Error>;
	}
	export interface NextLink<T> {
	    (callback?: PromiseCallback<T>): Promise<T, Error>;
	}
	export class Users extends DataModelListWrapper<User, Users> {
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
	    attachments: AttachmentDataModel[];
	    bccRecipients: Recipient[];
	    body: ItemBody;
	    bodyPreview: string;
	    categories: string[];
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
	    replyTo: Recipient[];
	    sender: Recipient;
	    sentDateTime: string;
	    subject: string;
	    toRecipients: Recipient[];
	    webLink: string;
	}
	export class Message extends DataModelWrapper<MessageDataModel> {
	}
	export class Messages extends DataModelListWrapper<Message, Messages> {
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
	export interface PatternedRecurrence {
	}
	export interface ResponseStatus {
	    response: string;
	    time: string;
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
	export class Event extends DataModelWrapper<EventDataModel> {
	}
	export class Events extends DataModelListWrapper<Event, Events> {
	    protected graph: Graph;
	    protected endpoint: EventsEndpoint;
	    protected _data: Event[];
	    constructor(graph: Graph, endpoint: EventsEndpoint, _data: Event[]);
	}
	export class Contact {
	}
	export class GroupDataModel {
	    id: string;
	    description: string;
	    displayName: string;
	    groupTypes: string[];
	    mail: string;
	    mailEnabled: Boolean;
	    mailNickname: string;
	    onPremisesLastSyncDateTime: Date;
	    onPremisesSecurityIdentifier: string;
	    onPremisesSyncEnabled: Boolean;
	    proxyAddresses: string[];
	    securityEnabled: Boolean;
	    visibility: string;
	}
	export class Group extends DataModelWrapper<GroupDataModel> {
	}
	export class Groups extends DataModelListWrapper<Group, Groups> {
	}
	export class MailFolderDataModel {
	    id: string;
	    displayName: string;
	    childFolderCount: number;
	    unreadItemCount: number;
	    totalItemCount: number;
	}
	export class MailFolder extends DataModelWrapper<MailFolderDataModel> {
	}
	export class MailFolders extends DataModelListWrapper<MailFolder, MailFolders> {
	}
	export enum AttachmentType {
	    fileAttachment = 0,
	    itemAttachment = 1,
	    referenceAttachment = 2,
	}
	export class AttachmentDataModel {
	    contentId: string;
	    id: string;
	    isInline: boolean;
	    lastModifiedDateTime: Date;
	    name: string;
	    size: number;
	    contentBytes: string;
	    contentLocation: string;
	    contentType: string;
	}
	export class Attachment extends DataModelWrapper<AttachmentDataModel> {
	    getType(): AttachmentType;
	}
	export class Attachments extends DataModelListWrapper<Attachment, Attachments> {
	}
	export class Graph {
	    private req;
	    private accessToken;
	    private KurveIdentity;
	    private defaultResourceID;
	    private baseUrl;
	    constructor(identityInfo: {
	        identity: Identity;
	    });
	    constructor(identityInfo: {
	        defaultAccessToken: string;
	    });
	    private scopesForV2(scopes);
	    meAsync(odataQuery?: string): Promise<User, Error>;
	    me(callback: PromiseCallback<User>, odataQuery?: string): void;
	    userAsync(userId: string, odataQuery?: string, basicProfileOnly?: boolean): Promise<User, Error>;
	    user(userId: string, callback: PromiseCallback<User>, odataQuery?: string, basicProfileOnly?: boolean): void;
	    usersAsync(odataQuery?: string, basicProfileOnly?: boolean): Promise<Users, Error>;
	    users(callback: PromiseCallback<Users>, odataQuery?: string, basicProfileOnly?: boolean): void;
	    groupAsync(groupId: string, odataQuery?: string): Promise<Group, Error>;
	    group(groupId: string, callback: PromiseCallback<Group>, odataQuery?: string): void;
	    groupsAsync(odataQuery?: string): Promise<Groups, Error>;
	    groups(callback: PromiseCallback<Groups>, odataQuery?: string): void;
	    messageForUserAsync(userPrincipalName: string, messageId: string, odataQuery?: string): Promise<Message, Error>;
	    messageForUser(userPrincipalName: string, messageId: string, callback: PromiseCallback<Message>, odataQuery?: string): void;
	    messagesForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error>;
	    messagesForUser(userPrincipalName: string, callback: PromiseCallback<Messages>, odataQuery?: string): void;
	    mailFoldersForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<MailFolders, Error>;
	    mailFoldersForUser(userPrincipalName: string, callback: PromiseCallback<MailFolders>, odataQuery?: string): void;
	    eventForUserAsync(userPrincipalName: string, eventId: string, odataQuery?: string): Promise<Event, Error>;
	    eventForUser(userPrincipalName: string, eventId: string, callback: PromiseCallback<Event>, odataQuery?: string): void;
	    eventsForUserAsync(userPrincipalName: string, endpoint: EventsEndpoint, odataQuery?: string): Promise<Events, Error>;
	    eventsForUser(userPrincipalName: string, endpoint: EventsEndpoint, callback: (messages: Events, error: Error) => void, odataQuery?: string): void;
	    memberOfForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Groups, Error>;
	    memberOfForUser(userPrincipalName: string, callback: PromiseCallback<Groups>, odataQuery?: string): void;
	    managerForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<User, Error>;
	    managerForUser(userPrincipalName: string, callback: PromiseCallback<User>, odataQuery?: string): void;
	    directReportsForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Users, Error>;
	    directReportsForUser(userPrincipalName: string, callback: PromiseCallback<Users>, odataQuery?: string): void;
	    profilePhotoForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<ProfilePhoto, Error>;
	    profilePhotoForUser(userPrincipalName: string, callback: PromiseCallback<ProfilePhoto>, odataQuery?: string): void;
	    profilePhotoValueForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<any, Error>;
	    profilePhotoValueForUser(userPrincipalName: string, callback: PromiseCallback<any>, odataQuery?: string): void;
	    messageAttachmentsForUserAsync(userPrincipalName: string, messageId: string, odataQuery?: string): Promise<Attachments, Error>;
	    messageAttachmentsForUser(userPrincipalName: string, messageId: string, callback: PromiseCallback<Attachments>, odataQuery?: string): void;
	    messageAttachmentForUserAsync(userPrincipalName: string, messageId: string, attachmentId: string, odataQuery?: string): Promise<Attachment, Error>;
	    messageAttachmentForUser(userPrincipalName: string, messageId: string, attachmentId: string, callback: PromiseCallback<Attachment>, odataQuery?: string): void;
	    getAsync(url: string): Promise<string, Error>;
	    get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?: string[]): void;
	    private generateError(xhr);
	    private getUsers(urlString, callback, scopes?, basicProfileOnly?);
	    private getUser(urlString, callback, scopes?);
	    private addAccessTokenAndSend(xhr, callback, scopes?);
	    private getMessage(urlString, messageId, callback, scopes?);
	    private getMessages(urlString, callback, scopes?);
	    private getEvent(urlString, EventId, callback, scopes?);
	    private getEvents(urlString, endpoint, callback, scopes?);
	    private getGroups(urlString, callback, scopes?);
	    private getGroup(urlString, callback, scopes?);
	    private getPhoto(urlString, callback, scopes?);
	    private getPhotoValue(urlString, callback, scopes?);
	    private getMailFolders(urlString, callback, scopes?);
	    private getMessageAttachments(urlString, callback, scopes?);
	    private getMessageAttachment(urlString, callback, scopes?);
	    private buildUrl(root, path, odataQuery?);
	    private buildMeUrl(path?, odataQuery?);
	    private buildUsersUrl(path?, odataQuery?);
	    private buildGroupsUrl(path?, odataQuery?);
	}

}
declare module 'Kurve/src/kurve' {
	export * from 'Kurve/src/graph';
	export * from 'Kurve/src/identity';

}
declare module 'Kurve' {
	import main = require('Kurve/src/kurve');
	export = main;
}
