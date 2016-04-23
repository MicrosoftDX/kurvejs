declare module Kurve {
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
    class Error {
        status: number;
        statusText: string;
        text: string;
        other: any;
    }
    class Identity {
        authContext: any;
        config: any;
        isCallback: boolean;
        clientId: string;
        private req;
        private state;
        private nonce;
        private idToken;
        private loginCallback;
        private accessTokenCallback;
        private getTokenCallback;
        private redirectUri;
        private tokenCache;
        private logonUser;
        private refreshTimer;
        constructor(clientId?: string, redirectUri?: string);
        private decodeIdToken(idToken);
        private decodeAccessToken(accessToken, resource);
        getIdToken(): any;
        isLoggedIn(): boolean;
        private renewIdToken();
        getAccessTokenAsync(resource: string): Promise<string, Error>;
        getAccessToken(resource: string, callback: (token: string, error: Error) => void): void;
        loginAsync(): Promise<void, Error>;
        login(callback: (error: Error) => void): void;
        logOut(): void;
        private base64Decode(encodedString);
        private generateNonce();
    }
}
declare module Kurve {
    class ProfilePhotoDataModel {
        id: string;
        height: Number;
        width: Number;
    }
    class ProfilePhoto {
        protected graph: Kurve.Graph;
        protected _data: ProfilePhotoDataModel;
        constructor(graph: Kurve.Graph, _data: ProfilePhotoDataModel);
        data: ProfilePhotoDataModel;
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
    class User {
        protected graph: Kurve.Graph;
        protected _data: UserDataModel;
        constructor(graph: Kurve.Graph, _data: UserDataModel);
        data: UserDataModel;
        memberOf(callback: (groups: Groups, Error) => void, Error: any, odataQuery?: string): void;
        memberOfAsync(odataQuery?: string): Promise<Messages, Error>;
        messages(callback: (messages: Messages, error: Error) => void, odataQuery?: string): void;
        messagesAsync(odataQuery?: string): Promise<Messages, Error>;
        manager(callback: (user: Kurve.User, error: Error) => void, odataQuery?: string): void;
        managerAsync(odataQuery?: string): Promise<User, Error>;
        profilePhoto(callback: (photo: ProfilePhoto, error: Error) => void): void;
        profilePhotoAsync(): Promise<ProfilePhoto, Error>;
        profilePhotoValue(callback: (val: any, error: Error) => void): void;
        profilePhotoValueAsync(): Promise<any, Error>;
        calendar(callback: (calendarItems: CalendarEvents, error: Error) => void, odataQuery?: string): void;
        calendarAsync(odataQuery?: string): Promise<CalendarEvents, Error>;
    }
    class Users {
        protected graph: Kurve.Graph;
        protected _data: User[];
        nextLink: (callback?: (users: Kurve.Users, error: Error) => void, odataQuery?: string) => Promise<Users, Error>;
        constructor(graph: Kurve.Graph, _data: User[]);
        data: User[];
    }
    class MessageDataModel {
        bccRecipients: string[];
        body: Object;
        bodyPreview: string;
        categories: string[];
        ccRecipients: string[];
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
        replyTo: any[];
        sender: any;
        sentDateTime: string;
        subject: string;
        toRecipients: string[];
        webLink: string;
    }
    class Message {
        protected graph: Kurve.Graph;
        protected _data: MessageDataModel;
        constructor(graph: Kurve.Graph, _data: MessageDataModel);
        data: MessageDataModel;
    }
    class Messages {
        protected graph: Kurve.Graph;
        protected _data: Message[];
        nextLink: (callback?: (messages: Kurve.Messages, error: Error) => void, odataQuery?: string) => Promise<Messages, Error>;
        constructor(graph: Kurve.Graph, _data: Message[]);
        data: Message[];
    }
    class CalendarEvent {
    }
    class CalendarEvents {
        protected graph: Kurve.Graph;
        protected _data: CalendarEvent[];
        nextLink: (callback?: (events: Kurve.CalendarEvents, error: Error) => void, odataQuery?: string) => Promise<(events: Kurve.CalendarEvents, error: Error) => void, Error>;
        constructor(graph: Kurve.Graph, _data: CalendarEvent[]);
        data: CalendarEvent[];
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
    class Group {
        protected graph: Kurve.Graph;
        protected _data: GroupDataModel;
        constructor(graph: Kurve.Graph, _data: GroupDataModel);
        data: GroupDataModel;
    }
    class Groups {
        protected graph: Kurve.Graph;
        protected _data: Group[];
        nextLink: (callback?: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) => Promise<Groups, Error>;
        constructor(graph: Kurve.Graph, _data: Group[]);
        data: Group[];
    }
    class Graph {
        private req;
        private state;
        private nonce;
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
        meAsync(odataQuery?: string): Promise<User, Error>;
        me(callback: (user: User, error: Error) => void, odataQuery?: string): void;
        userAsync(userId: string): Promise<User, Error>;
        user(userId: string, callback: (user: Kurve.User, error: Error) => void): void;
        usersAsync(odataQuery?: string): Promise<Users, Error>;
        users(callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string): void;
        groupAsync(groupId: string): Promise<Group, Error>;
        group(groupId: string, callback: (group: any, error: Error) => void): void;
        groups(callback: (groups: any, error: Error) => void, odataQuery?: string): void;
        groupsAsync(odataQuery?: string): Promise<Groups, Error>;
        messagesForUser(userPrincipalName: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void;
        messagesForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error>;
        calendarForUser(userPrincipalName: string, callback: (events: CalendarEvent, error: Error) => void, odataQuery?: string): void;
        calendarForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<CalendarEvents, Error>;
        memberOfForUser(userPrincipalName: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string): void;
        memberOfForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error>;
        managerForUser(userPrincipalName: string, callback: (manager: Kurve.User, error: Error) => void, odataQuery?: string): void;
        managerForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<User, Error>;
        directReportsForUser(userPrincipalName: string, callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string): void;
        directReportsForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Users, Error>;
        profilePhotoForUser(userPrincipalName: string, callback: (photo: ProfilePhoto, error: Error) => void): void;
        profilePhotoForUserAsync(userPrincipalName: string): Promise<ProfilePhoto, Error>;
        profilePhotoValueForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void): void;
        profilePhotoValueForUserAsync(userPrincipalName: string): Promise<any, Error>;
        getAsync(url: string): Promise<string, Error>;
        get(url: string, callback: (response: string, error: Error) => void, responseType?: string): void;
        private generateError(xhr);
        private getUsers(urlString, callback);
        private getUser(urlString, callback);
        private addAccessTokenAndSend(xhr, callback);
        private getMessages(urlString, callback, odataQuery?);
        private getGroups(urlString, callback, odataQuery?);
        private getGroup(urlString, callback);
        private getPhoto(urlString, callback);
        private getPhotoValue(urlString, callback);
        private buildMeUrl();
        private buildUsersUrl();
        private buildGroupsUrl();
    }
}
