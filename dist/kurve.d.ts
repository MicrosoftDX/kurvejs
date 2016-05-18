declare namespace __Kurve {
    class Error {
        status: number;
        statusText: string;
        text: string;
        other: any;
    }
    interface PromiseCallback<T> {
        (error: Error, result?: T): void;
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
    class Promise<T, E> {
        private _deferred;
        constructor(_deferred: Deferred<T, E>);
        then<R>(successCallback?: (result: T) => R, errorCallback?: (error: E) => R): any;
        fail<R>(errorCallback?: (error: E) => R): any;
    }
}
declare namespace __Kurve {
    enum EndPointVersion {
        v1 = 1,
        v2 = 2,
    }
    enum Mode {
        Client = 1,
        Node = 2,
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
        appSecret?: string;
        tokenProcessingUri: string;
        version: EndPointVersion;
        tokenStorage?: TokenStorage;
        mode: Mode;
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
        private mode;
        private appSecret;
        private NodePersistDataCallBack;
        private NodeRetrieveDataCallBack;
        private req;
        private res;
        private https;
        constructor(identitySettings: IdentitySettings);
        private parseQueryString(str);
        private token(s, url);
        checkForIdentityRedirect(): boolean;
        private decodeIdToken(idToken);
        private decodeAccessToken(accessToken, resource?, scopes?);
        getIdToken(): any;
        isLoggedIn(): boolean;
        private renewIdToken();
        getCurrentEndPointVersion(): EndPointVersion;
        getAccessTokenAsync(resource: string): Promise<string, Error>;
        getAccessToken(resource: string, callback: PromiseCallback<string>): void;
        private parseNodeCookies(req);
        handleNodeCallback(req: any, res: any, https: any, crypto: any, persistDataCallback: (key: string, value: string, expiry: Date) => void, retrieveDataCallback: (key: string) => string): Promise<boolean, Error>;
        getAccessTokenForScopesAsync(scopes: string[], promptForConsent?: boolean): Promise<string, Error>;
        getAccessTokenForScopes(scopes: string[], promptForConsent: any, callback: (token: string, error: Error) => void): void;
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
declare namespace __Kurve {
    class Graph {
        private req;
        private accessToken;
        KurveIdentity: Identity;
        private defaultResourceID;
        private baseUrl;
        private https;
        private mode;
        constructor(identityInfo: {
            identity: Identity;
        }, mode: Mode, https?: any);
        constructor(identityInfo: {
            defaultAccessToken: string;
        }, mode: Mode, https?: any);
        me: User;
        users: Users;
        groups: Groups;
        get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?: string[]): void;
        private findAccessToken(callback, scopes?);
        post(object: string, url: string, callback: PromiseCallback<string>, responseType?: string, scopes?: string[]): void;
        private generateError(xhr);
        private addAccessTokenAndSend(xhr, callback, scopes?);
    }
}
declare namespace __Kurve {
    interface ItemBody {
        contentType?: string;
        content?: string;
    }
    interface EmailAddress {
        name?: string;
        address?: string;
    }
    interface Recipient {
        emailAddress?: EmailAddress;
    }
    interface UserDataModel {
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
    interface ProfilePhotoDataModel {
        id?: string;
        height?: Number;
        width?: Number;
    }
    interface MessageDataModel {
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
    interface Attendee {
        status?: ResponseStatus;
        type?: string;
        emailAddress?: EmailAddress;
    }
    interface DateTimeTimeZone {
        dateTime?: string;
        timeZone?: string;
    }
    interface PatternedRecurrence {
    }
    interface ResponseStatus {
        response?: string;
        time?: string;
    }
    interface Location {
        displayName?: string;
        address?: any;
    }
    interface EventDataModel {
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
    interface GroupDataModel {
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
    interface MailFolderDataModel {
        id?: string;
        displayName?: string;
        childFolderCount?: number;
        unreadItemCount?: number;
        totalItemCount?: number;
    }
    interface AttachmentDataModel {
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
declare namespace __Kurve {
    class Scopes {
        private static rootUrl;
        static General: {
            OpenId: string;
            OfflineAccess: string;
        };
        static User: {
            Read: string;
            ReadAll: string;
            ReadWrite: string;
            ReadWriteAll: string;
            ReadBasicAll: string;
        };
        static Contacts: {
            Read: string;
            ReadWrite: string;
        };
        static Directory: {
            ReadAll: string;
            ReadWriteAll: string;
            AccessAsUserAll: string;
        };
        static Group: {
            ReadAll: string;
            ReadWriteAll: string;
            AccessAsUserAll: string;
        };
        static Mail: {
            Read: string;
            ReadWrite: string;
            Send: string;
        };
        static Calendars: {
            Read: string;
            ReadWrite: string;
        };
        static Files: {
            Read: string;
            ReadAll: string;
            ReadWrite: string;
            ReadWriteAppFolder: string;
            ReadWriteSelected: string;
        };
        static Tasks: {
            ReadWrite: string;
        };
        static People: {
            Read: string;
            ReadWrite: string;
        };
        static Notes: {
            Create: string;
            ReadWriteCreatedByApp: string;
            Read: string;
            ReadAll: string;
            ReadWriteAll: string;
        };
    }
    class OData {
        query: string;
        constructor(query?: string);
        toString: () => string;
        odata: (query: string) => this;
        select: (...fields: string[]) => this;
        expand: (...fields: string[]) => this;
        filter: (query: string) => this;
        orderby: (...fields: string[]) => this;
        top: (items: Number) => this;
        skip: (items: Number) => this;
    }
    type GraphObject<Model, N extends Node> = Model & {
        _node?: N;
        _item: Model;
    };
    type ChildFactory<Model, N extends Node> = (id: string) => N;
    type GraphCollection<Model, C extends CollectionNode, N extends Node> = Array<GraphObject<Model, N>> & {
        _next?: () => Promise<GraphCollection<Model, C, N>, Error>;
        _node?: C;
        _raw: any;
        _items: Model[];
    };
    abstract class Node {
        protected graph: Graph;
        protected path: string;
        constructor(graph: Graph, path: string);
        protected scopesForV2: (scopes: string[]) => string[];
        pathWithQuery: (odataQuery?: OData | string, pathSuffix?: string) => string;
        protected graphObjectFromResponse: <Model, N extends Node>(response: any, node: N) => Model & {
            _node?: N;
            _item: Model;
        };
        protected get<Model, N extends Node>(path: string, node: N, scopes?: string[], responseType?: string): Promise<GraphObject<Model, N>, Error>;
        protected post<Model, N extends Node>(object: Model, path: string, node: N, scopes?: string[]): Promise<GraphObject<Model, N>, Error>;
    }
    abstract class CollectionNode extends Node {
        private _nextLink;
        pathWithQuery: (odataQuery?: OData | string, pathSuffix?: string) => string;
        nextLink: string;
        protected graphCollectionFromResponse: <Model, C extends CollectionNode, N extends Node>(response: any, node: C, childFactory?: (id: string) => N, scopes?: string[]) => (Model & {
            _node?: N;
            _item: Model;
        })[] & {
            _next?: () => Promise<(Model & {
                _node?: N;
                _item: Model;
            })[] & any, Error>;
            _node?: C;
            _raw: any;
            _items: Model[];
        };
        protected getCollection<Model, C extends CollectionNode, N extends Node>(path: string, node: C, childFactory: ChildFactory<Model, N>, scopes?: string[]): Promise<GraphCollection<Model, C, N>, Error>;
    }
    class Attachment extends Node {
        private context;
        constructor(graph: Graph, path: string, context: string, attachmentId?: string);
        static scopes: {
            messages: string[];
            events: string[];
        };
        GetAttachment: (odataQuery?: OData | string) => Promise<AttachmentDataModel & {
            _node?: Attachment;
            _item: AttachmentDataModel;
        }, Error>;
    }
    class Attachments extends CollectionNode {
        private context;
        constructor(graph: Graph, path: string, context: string);
        $: (attachmentId: string) => Attachment;
        GetAttachments: (odataQuery?: OData | string) => Promise<(AttachmentDataModel & {
            _node?: Attachment;
            _item: AttachmentDataModel;
        })[] & {
            _next?: () => Promise<(AttachmentDataModel & {
                _node?: Attachment;
                _item: AttachmentDataModel;
            })[] & any, Error>;
            _node?: Attachments;
            _raw: any;
            _items: AttachmentDataModel[];
        }, Error>;
    }
    class Message extends Node {
        constructor(graph: Graph, path?: string, messageId?: string);
        attachments: Attachments;
        GetMessage: (odataQuery?: OData | string) => Promise<MessageDataModel & {
            _node?: Message;
            _item: MessageDataModel;
        }, Error>;
        SendMessage: (odataQuery?: OData | string) => Promise<MessageDataModel & {
            _node?: Message;
            _item: MessageDataModel;
        }, Error>;
    }
    class Messages extends CollectionNode {
        constructor(graph: Graph, path?: string);
        $: (messageId: string) => Message;
        GetMessages: (odataQuery?: OData | string) => Promise<(MessageDataModel & {
            _node?: Message;
            _item: MessageDataModel;
        })[] & {
            _next?: () => Promise<(MessageDataModel & {
                _node?: Message;
                _item: MessageDataModel;
            })[] & any, Error>;
            _node?: Messages;
            _raw: any;
            _items: MessageDataModel[];
        }, Error>;
        CreateMessage: (object: MessageDataModel, odataQuery?: OData | string) => Promise<MessageDataModel & {
            _node?: Messages;
            _item: MessageDataModel;
        }, Error>;
    }
    class Event extends Node {
        constructor(graph: Graph, path: string, eventId: string);
        attachments: Attachments;
        GetEvent: (odataQuery?: OData | string) => Promise<EventDataModel & {
            _node?: Event;
            _item: EventDataModel;
        }, Error>;
    }
    class Events extends CollectionNode {
        constructor(graph: Graph, path?: string);
        $: (eventId: string) => Event;
        GetEvents: (odataQuery?: OData | string) => Promise<(EventDataModel & {
            _node?: Event;
            _item: EventDataModel;
        })[] & {
            _next?: () => Promise<(EventDataModel & {
                _node?: Event;
                _item: EventDataModel;
            })[] & any, Error>;
            _node?: Events;
            _raw: any;
            _items: EventDataModel[];
        }, Error>;
    }
    class CalendarView extends CollectionNode {
        private static suffix;
        constructor(graph: Graph, path?: string);
        private $;
        dateRange: (startDate: Date, endDate: Date) => string;
        GetCalendarView: (odataQuery?: OData | string) => Promise<(EventDataModel & {
            _node?: Event;
            _item: EventDataModel;
        })[] & {
            _next?: () => Promise<(EventDataModel & {
                _node?: Event;
                _item: EventDataModel;
            })[] & any, Error>;
            _node?: CalendarView;
            _raw: any;
            _items: EventDataModel[];
        }, Error>;
    }
    class MailFolder extends Node {
        constructor(graph: Graph, path: string, mailFolderId: string);
        GetMailFolder: (odataQuery?: OData | string) => Promise<MailFolderDataModel & {
            _node?: MailFolder;
            _item: MailFolderDataModel;
        }, Error>;
    }
    class MailFolders extends CollectionNode {
        constructor(graph: Graph, path?: string);
        $: (mailFolderId: string) => MailFolder;
        GetMailFolders: (odataQuery?: OData | string) => Promise<(MailFolderDataModel & {
            _node?: MailFolder;
            _item: MailFolderDataModel;
        })[] & {
            _next?: () => Promise<(MailFolderDataModel & {
                _node?: MailFolder;
                _item: MailFolderDataModel;
            })[] & any, Error>;
            _node?: MailFolders;
            _raw: any;
            _items: MailFolderDataModel[];
        }, Error>;
    }
    class Photo extends Node {
        private context;
        constructor(graph: Graph, path: string, context: string);
        static scopes: {
            user: string[];
            group: string[];
            contact: string[];
        };
        GetPhotoProperties: (odataQuery?: OData | string) => Promise<ProfilePhotoDataModel & {
            _node?: Photo;
            _item: ProfilePhotoDataModel;
        }, Error>;
        GetPhotoImage: (odataQuery?: OData | string) => Promise<any, Error>;
    }
    class Manager extends Node {
        constructor(graph: Graph, path?: string);
        GetManager: (odataQuery?: OData | string) => Promise<UserDataModel & {
            _node?: Manager;
            _item: UserDataModel;
        }, Error>;
    }
    class MemberOf extends CollectionNode {
        constructor(graph: Graph, path?: string);
        GetGroups: (odataQuery?: OData | string) => Promise<(GroupDataModel & {
            _node?: Group;
            _item: GroupDataModel;
        })[] & {
            _next?: () => Promise<(GroupDataModel & {
                _node?: Group;
                _item: GroupDataModel;
            })[] & any, Error>;
            _node?: MemberOf;
            _raw: any;
            _items: GroupDataModel[];
        }, Error>;
    }
    class DirectReport extends Node {
        protected graph: Graph;
        constructor(graph: Graph, path?: string, userId?: string);
        GetDirectReport: (odataQuery?: OData | string) => Promise<UserDataModel & {
            _node?: DirectReport;
            _item: UserDataModel;
        }, Error>;
    }
    class DirectReports extends CollectionNode {
        constructor(graph: Graph, path?: string);
        $: (userId: string) => DirectReport;
        GetDirectReports: (odataQuery?: OData | string) => Promise<(UserDataModel & {
            _node?: User;
            _item: UserDataModel;
        })[] & {
            _next?: () => Promise<(UserDataModel & {
                _node?: User;
                _item: UserDataModel;
            })[] & any, Error>;
            _node?: DirectReports;
            _raw: any;
            _items: UserDataModel[];
        }, Error>;
    }
    class User extends Node {
        protected graph: Graph;
        constructor(graph: Graph, path?: string, userId?: string);
        messages: Messages;
        events: Events;
        calendarView: CalendarView;
        mailFolders: MailFolders;
        photo: Photo;
        manager: Manager;
        directReports: DirectReports;
        memberOf: MemberOf;
        GetUser: (odataQuery?: OData | string) => Promise<UserDataModel & {
            _node?: User;
            _item: UserDataModel;
        }, Error>;
    }
    class Users extends CollectionNode {
        constructor(graph: Graph, path?: string);
        $: (userId: string) => User;
        static $: (graph: Graph) => (userId: string) => User;
        GetUsers: (odataQuery?: OData | string) => Promise<(UserDataModel & {
            _node?: User;
            _item: UserDataModel;
        })[] & {
            _next?: () => Promise<(UserDataModel & {
                _node?: User;
                _item: UserDataModel;
            })[] & any, Error>;
            _node?: Users;
            _raw: any;
            _items: UserDataModel[];
        }, Error>;
    }
    class Group extends Node {
        protected graph: Graph;
        constructor(graph: Graph, path: string, groupId: string);
        GetGroup: (odataQuery?: OData | string) => Promise<GroupDataModel & {
            _node?: Group;
            _item: GroupDataModel;
        }, Error>;
    }
    class Groups extends CollectionNode {
        constructor(graph: Graph, path?: string);
        $: (groupId: string) => Group;
        static $: (graph: Graph) => (groupId: string) => Group;
        GetGroups: (odataQuery?: OData | string) => Promise<(GroupDataModel & {
            _node?: Group;
            _item: GroupDataModel;
        })[] & {
            _next?: () => Promise<(GroupDataModel & {
                _node?: Group;
                _item: GroupDataModel;
            })[] & any, Error>;
            _node?: Groups;
            _raw: any;
            _items: GroupDataModel[];
        }, Error>;
    }
}
declare module "kurve" {
    export = __Kurve;
}
