declare module Kurve {
    interface PromiseCallback<T> {
        (result: T, error: Error): void;
    }
    class Deferred<T, E> {
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
    class Promise<T, E> implements Promise<T, E> {
        private _deferred;
        constructor(_deferred: Deferred<T, E>);
        then<R>(successCallback?: (result: T) => R, errorCallback?: (error: E) => R): Promise<R, E>;
        fail<R>(errorCallback?: (error: E) => R): Promise<R, E>;
    }
}
declare module Kurve {
    enum OAuthVersion {
        v1 = 1,
        v2 = 2,
    }
    class Error {
        status: number;
        statusText: string;
        text: string;
        other: any;
    }
    interface TokenStorage {
        add(key: string, token: any): any;
        remove(key: string): any;
        getAll(): any[];
        clear(): any;
    }
    class IdToken {
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
    interface IdentitySettings {
        clientId: string;
        tokenProcessingUri: string;
        version: OAuthVersion;
        tokenStorage?: TokenStorage;
    }
    class Identity {
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
declare module Kurve {
    module Scopes {
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
    class DataModelWrapper<T> {
        protected graph: Graph;
        protected _data: T;
        constructor(graph: Graph, _data: T);
        data: T;
    }
    class DataModelListWrapper<T, S> extends DataModelWrapper<T[]> {
        nextLink: NextLink<S>;
    }
    class ProfilePhotoDataModel {
        id: string;
        height: Number;
        width: Number;
    }
    class ProfilePhoto extends DataModelWrapper<ProfilePhotoDataModel> {
    }
    class UserDataModel {
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
    enum EventsEndpoint {
        events = 0,
        calendarView = 1,
    }
    class User extends DataModelWrapper<UserDataModel> {
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
        event(eventId: string, callback: PromiseCallback<Event>, odataQuery?: string): void;
        eventAsync(eventId: string, odataQuery?: string): Promise<Event, Error>;
        messageAttachment(messageId: string, attachmentId: string, callback: PromiseCallback<Attachment>, odataQuery?: string): void;
        messageAttachmentAsync(messageId: string, attachmentId: string, odataQuery?: string): Promise<Attachment, Error>;
    }
    interface NextLink<T> {
        (callback?: PromiseCallback<T>): Promise<T, Error>;
    }
    class Users extends DataModelListWrapper<User, Users> {
    }
    interface ItemBody {
        contentType: string;
        content: string;
    }
    interface EmailAddress {
        name: string;
        address: string;
    }
    interface Recipient {
        emailAddress: EmailAddress;
    }
    class MessageDataModel {
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
    class Message extends DataModelWrapper<MessageDataModel> {
    }
    class Messages extends DataModelListWrapper<Message, Messages> {
    }
    interface Attendee {
        status: ResponseStatus;
        type: string;
        emailAddress: EmailAddress;
    }
    interface DateTimeTimeZone {
        dateTime: string;
        timeZone: string;
    }
    interface PatternedRecurrence {
    }
    interface ResponseStatus {
        response: string;
        time: string;
    }
    interface Location {
        displayName: string;
        address: any;
    }
    class EventDataModel {
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
    class Event extends DataModelWrapper<EventDataModel> {
    }
    class Events extends DataModelListWrapper<Event, Events> {
        protected graph: Graph;
        protected endpoint: EventsEndpoint;
        protected _data: Event[];
        constructor(graph: Graph, endpoint: EventsEndpoint, _data: Event[]);
    }
    class Contact {
    }
    class GroupDataModel {
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
    class Group extends DataModelWrapper<GroupDataModel> {
    }
    class Groups extends DataModelListWrapper<Group, Groups> {
    }
    class MailFolderDataModel {
        id: string;
        displayName: string;
        childFolderCount: number;
        unreadItemCount: number;
        totalItemCount: number;
    }
    class MailFolder extends DataModelWrapper<MailFolderDataModel> {
    }
    class MailFolders extends DataModelListWrapper<MailFolder, MailFolders> {
    }
    enum AttachmentType {
        fileAttachment = 0,
        itemAttachment = 1,
        referenceAttachment = 2,
    }
    class AttachmentDataModel {
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
    class Attachment extends DataModelWrapper<AttachmentDataModel> {
        getType(): AttachmentType;
    }
    class Attachments extends DataModelListWrapper<Attachment, Attachments> {
    }
    class Graph {
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
        scopesForV2(scopes: string[]): string[];
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
    class GraphInfo<T> {
        protected path: string;
        protected scopes: string[];
        protected odataQuery: string;
        constructor(path: string, scopes: string[], odataQuery?: string);
        getAsync(graph: Graph): Promise<T, Error>;
        get(graph: Graph, callback: PromiseCallback<T>): void;
    }
    class ItemAttachmentDataModel {
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
    class ItemAttachment extends ItemAttachmentDataModel {
        static fromMessageForMe: (messageId: string, attachmentId: string, odataQuery?: string) => GraphInfo<ItemAttachmentDataModel>;
        static fromMessageForUser: (userId: string, messageId: string, attachmentId: string, odataQuery?: string) => GraphInfo<ItemAttachmentDataModel>;
    }
}
