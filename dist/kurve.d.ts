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
	import { Error } from 'Kurve/src/identity';
	export interface PromiseCallback<T> {
	    (error: Error, result?: T): void;
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
declare module 'Kurve/src/models' {
	export interface ItemBody {
	    contentType?: string;
	    content?: string;
	}
	export interface EmailAddress {
	    name?: string;
	    address?: string;
	}
	export interface Recipient {
	    emailAddress?: EmailAddress;
	}
	export interface UserDataModel {
	    businessPhones?: string;
	    displayName?: string;
	    givenName?: string;
	    jobTitle?: string;
	    mail?: string;
	    mobilePhone?: string;
	    officeLocation?: string;
	    preferredLanguage?: string;
	    surname?: string;
	    userPrincipalName?: string;
	    id?: string;
	}
	export interface ProfilePhotoDataModel {
	    id?: string;
	    height?: Number;
	    width?: Number;
	}
	export interface MessageDataModel {
	    attachments?: AttachmentDataModel[];
	    bccRecipients?: Recipient[];
	    body?: ItemBody;
	    bodyPreview?: string;
	    categories?: string[];
	    ccRecipients?: Recipient[];
	    changeKey?: string;
	    conversationId?: string;
	    createdDateTime?: string;
	    from?: Recipient;
	    graph?: any;
	    hasAttachments?: boolean;
	    id?: string;
	    importance?: string;
	    isDeliveryReceiptRequested?: boolean;
	    isDraft?: boolean;
	    isRead?: boolean;
	    isReadReceiptRequested?: boolean;
	    lastModifiedDateTime?: string;
	    parentFolderId?: string;
	    receivedDateTime?: string;
	    replyTo?: Recipient[];
	    sender?: Recipient;
	    sentDateTime?: string;
	    subject?: string;
	    toRecipients?: Recipient[];
	    webLink?: string;
	}
	export interface Attendee {
	    status?: ResponseStatus;
	    type?: string;
	    emailAddress?: EmailAddress;
	}
	export interface DateTimeTimeZone {
	    dateTime?: string;
	    timeZone?: string;
	}
	export interface PatternedRecurrence {
	}
	export interface ResponseStatus {
	    response?: string;
	    time?: string;
	}
	export interface Location {
	    displayName?: string;
	    address?: any;
	}
	export interface EventDataModel {
	    attendees?: Attendee[];
	    body?: ItemBody;
	    bodyPreview?: string;
	    categories?: string[];
	    changeKey?: string;
	    createdDateTime?: string;
	    end?: DateTimeTimeZone;
	    hasAttachments?: boolean;
	    iCalUId?: string;
	    id?: string;
	    IDBCursor?: string;
	    importance?: string;
	    isAllDay?: boolean;
	    isCancelled?: boolean;
	    isOrganizer?: boolean;
	    isReminderOn?: boolean;
	    lastModifiedDateTime?: string;
	    location?: Location;
	    organizer?: Recipient;
	    originalEndTimeZone?: string;
	    originalStartTimeZone?: string;
	    recurrence?: PatternedRecurrence;
	    reminderMinutesBeforeStart?: number;
	    responseRequested?: boolean;
	    responseStatus?: ResponseStatus;
	    sensitivity?: string;
	    seriesMasterId?: string;
	    showAs?: string;
	    start?: DateTimeTimeZone;
	    subject?: string;
	    type?: string;
	    webLink?: string;
	}
	export interface GroupDataModel {
	    id?: string;
	    description?: string;
	    displayName?: string;
	    groupTypes?: string[];
	    mail?: string;
	    mailEnabled?: Boolean;
	    mailNickname?: string;
	    onPremisesLastSyncDateTime?: Date;
	    onPremisesSecurityIdentifier?: string;
	    onPremisesSyncEnabled?: Boolean;
	    proxyAddresses?: string[];
	    securityEnabled?: Boolean;
	    visibility?: string;
	}
	export interface MailFolderDataModel {
	    id?: string;
	    displayName?: string;
	    childFolderCount?: number;
	    unreadItemCount?: number;
	    totalItemCount?: number;
	}
	export interface AttachmentDataModel {
	    contentId?: string;
	    id?: string;
	    isInline?: boolean;
	    lastModifiedDateTime?: Date;
	    name?: string;
	    size?: number;
	    contentBytes?: string;
	    contentLocation?: string;
	    contentType?: string;
	}

}
declare module 'Kurve/src/requestbuilder' {
	import { Promise } from 'Kurve/src/promises';
	import { Graph } from 'Kurve/src/graph';
	import { Error } from 'Kurve/src/identity';
	import { UserDataModel, AttachmentDataModel, MessageDataModel, EventDataModel, MailFolderDataModel } from 'Kurve/src/models';
	export class Singleton<Model, N extends Node> {
	    raw: any;
	    self: N;
	    constructor(raw: any, self: N);
	    object: Model;
	}
	export class Collection<Model, N extends Node> {
	    raw: any;
	    self: N;
	    next: N;
	    constructor(raw: any, self: N, next: N);
	    objects: Model[];
	}
	export abstract class Node {
	    protected graph: Graph;
	    protected path: string;
	    protected query: string;
	    constructor(graph: Graph, path: string, query?: string);
	    pathWithQuery: string;
	    odata: (query: string) => this;
	    orderby: (...fields: string[]) => this;
	    top: (items: Number) => this;
	    skip: (items: Number) => this;
	    filter: (query: string) => this;
	    expand: (...fields: string[]) => this;
	    select: (...fields: string[]) => this;
	}
	export class Attachment extends Node {
	    constructor(graph: Graph, path: string, attachmentId: string);
	    GetAttachment: () => Promise<Singleton<AttachmentDataModel, Attachment>, Error>;
	}
	export class Attachments extends Node {
	    constructor(graph: Graph, path?: string);
	    id: (attachmentId: string) => Attachment;
	    GetAttachments: () => Promise<Collection<AttachmentDataModel, Attachments>, Error>;
	}
	export class Message extends Node {
	    constructor(graph: Graph, path: string, messageId: string);
	    attachments: () => Attachments;
	    GetMessage: () => Promise<Singleton<MessageDataModel, Message>, Error>;
	    SendMessage: () => Promise<Singleton<MessageDataModel, Message>, Error>;
	}
	export class Messages extends Node {
	    constructor(graph: Graph, path?: string);
	    id: (messageId: string) => Message;
	    GetMessages: () => Promise<Collection<MessageDataModel, Messages>, Error>;
	    CreateMessage: (object: MessageDataModel) => Promise<Singleton<MessageDataModel, Messages>, Error>;
	}
	export class Event extends Node {
	    constructor(graph: Graph, path: string, eventId: string);
	    attachments: () => Attachments;
	    GetEvent: () => Promise<Singleton<EventDataModel, Event>, Error>;
	}
	export class Events extends Node {
	    constructor(graph: Graph, path?: string);
	    id: (eventId: string) => Event;
	    GetEvents: () => Promise<Collection<EventDataModel, Events>, Error>;
	}
	export class CalendarView extends Node {
	    constructor(graph: Graph, path?: string);
	    dateRange: (startDate: Date, endDate: Date) => this;
	    GetCalendarView: () => Promise<Collection<EventDataModel, CalendarView>, Error>;
	}
	export class MailFolder extends Node {
	    constructor(graph: Graph, path: string, mailFolderId: string);
	    GetMailFolder: () => Promise<Singleton<MailFolderDataModel, MailFolder>, Error>;
	}
	export class MailFolders extends Node {
	    constructor(graph: Graph, path?: string);
	    id: (mailFolderId: string) => MailFolder;
	    GetMailFolders: () => Promise<Collection<MailFolderDataModel, MailFolders>, Error>;
	}
	export class User extends Node {
	    protected graph: Graph;
	    constructor(graph: Graph, path?: string, userId?: string);
	    messages: () => Messages;
	    events: () => Events;
	    calendarView: () => CalendarView;
	    mailFolders: () => MailFolders;
	    GetUser: () => Promise<Singleton<UserDataModel, User>, Error>;
	}
	export class Users extends Node {
	    constructor(graph: Graph, path?: string);
	    id: (userId: string) => User;
	    GetUsers: () => Promise<Collection<UserDataModel, Users>, Error>;
	}

}
declare module 'Kurve/src/graph' {
	import { Promise, PromiseCallback } from 'Kurve/src/promises';
	import { Identity, Error } from 'Kurve/src/identity';
	import { Singleton, Collection, Node, User, Users } from 'Kurve/src/requestbuilder';
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
	    me: () => User;
	    users: () => Users;
	    Get<Model, N extends Node>(path: string, self: N, scopes?: string[]): Promise<Singleton<Model, N>, Error>;
	    GetCollection<Model, N extends Node>(path: string, self: N, next: N, scopes?: string[]): Promise<Collection<Model, N>, Error>;
	    Post<Model, N extends Node>(object: Model, path: string, self: N, scopes?: string[]): Promise<Singleton<Model, N>, Error>;
	    private scopesForV2(scopes);
	    get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?: string[]): void;
	    post(object: string, url: string, callback: PromiseCallback<string>, responseType?: string, scopes?: string[]): void;
	    private generateError(xhr);
	    private addAccessTokenAndSend(xhr, callback, scopes?);
	}

}
declare module 'Kurve/src/kurve' {
	export * from 'Kurve/src/graph';
	export * from 'Kurve/src/identity';
	export * from 'Kurve/src/models';
	export * from 'Kurve/src/requestbuilder';

}
declare module 'Kurve' {
	import main = require('Kurve/src/kurve');
	export = main;
}
