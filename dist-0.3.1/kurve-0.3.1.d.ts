declare module Kurve {
    class Deferred {
        private doneCallbacks;
        private failCallbacks;
        private progressCallbacks;
        private _state;
        private _promise;
        private _result;
        private _notifyContext;
        private _notifyArgs;
        constructor();
        promise: Promise;
        state: string;
        rejected: boolean;
        resolved: boolean;
        resolve(...args: any[]): Deferred;
        resolveWith(context: any, ...args: any[]): Deferred;
        reject(...args: any[]): Deferred;
        rejectWith(context: any, ...args: any[]): Deferred;
        progress(...callbacks: Function[]): Deferred;
        notify(...args: any[]): Deferred;
        notifyWith(context: any, ...args: any[]): Deferred;
        private checkStatus();
        then(doneFilter: Function, failFilter?: Function, progressFilter?: Function): Deferred;
        private wrap(d, f, method);
        done(...callbacks: Function[]): Deferred;
        fail(...callbacks: Function[]): Deferred;
        always(...callbacks: Function[]): Deferred;
    }
    class Promise {
        protected deferred: Deferred;
        constructor(deferred: Deferred);
        then(doneFilter: Function, failFilter?: Function, progressFilter?: Function): Promise;
        done(...callbacks: Function[]): Promise;
        fail(...callbacks: Function[]): Promise;
        always(...callbacks: Function[]): Promise;
        resolved: boolean;
        rejected: boolean;
    }
    class TypedPromise<T> extends Promise {
        constructor(deferred: Deferred);
        then(doneFilter: (T) => void, failFilter?: Function, progressFilter?: Function): TypedPromise<T>;
        done(...callbacks: ((T) => void)[]): TypedPromise<T>;
        fail(...callbacks: Function[]): TypedPromise<T>;
        always(...callbacks: Function[]): TypedPromise<T>;
        resolved: boolean;
        rejected: boolean;
    }
    interface IWhen {
        (deferred: Deferred): Promise;
        (promise: Promise): Promise;
        (object: any): Promise;
        (...args: Deferred[]): Promise;
    }
    var when: IWhen;
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
        getAccessTokenAsync(resource: string): Promise;
        getAccessToken(resource: string, callback: (token: string, error: Error) => void): void;
        loginAsync(): Promise;
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
        memberOfAsync(odataQuery?: string): TypedPromise<Messages>;
        messages(callback: (messages: Messages, error: Error) => void, odataQuery?: string): void;
        messagesAsync(odataQuery?: string): TypedPromise<Messages>;
        manager(callback: (user: Kurve.User, error: Error) => void, odataQuery?: string): void;
        managerAsync(odataQuery?: string): TypedPromise<User>;
        profilePhoto(callback: (photo: ProfilePhoto, error: Error) => void): void;
        profilePhotoAsync(): TypedPromise<ProfilePhoto>;
        profilePhotoValue(callback: (val: any, error: Error) => void): void;
        profilePhotoValueAsync(): TypedPromise<any>;
        calendar(callback: (calendarItems: CalendarEvents, error: Error) => void, odataQuery?: string): void;
        calendarAsync(odataQuery?: string): TypedPromise<CalendarEvents>;
    }
    class Users {
        protected graph: Kurve.Graph;
        protected _data: User[];
        nextLink: (callback?: (users: Kurve.Users, error: Error) => void, odataQuery?: string) => Promise;
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
        nextLink: (callback?: (messages: Kurve.Messages, error: Error) => void, odataQuery?: string) => Promise;
        constructor(graph: Kurve.Graph, _data: Message[]);
        data: Message[];
    }
    class CalendarEvent {
    }
    class CalendarEvents {
        protected graph: Kurve.Graph;
        protected _data: CalendarEvent[];
        nextLink: (callback?: (events: Kurve.CalendarEvents, error: Error) => void, odataQuery?: string) => Promise;
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
        nextLink: (callback?: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) => TypedPromise<Groups>;
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
        meAsync(odataQuery?: string): TypedPromise<User>;
        me(callback: (user: User, error: Error) => void, odataQuery?: string): void;
        userAsync(userId: string): TypedPromise<User>;
        user(userId: string, callback: (user: Kurve.User, error: Error) => void): void;
        usersAsync(odataQuery?: string): TypedPromise<Users>;
        users(callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string): void;
        groupAsync(groupId: string): Promise;
        group(groupId: string, callback: (group: any, error: Error) => void): void;
        groups(callback: (groups: any, error: Error) => void, odataQuery?: string): void;
        groupsAsync(odataQuery?: string): Promise;
        messagesForUser(userPrincipalName: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void;
        messagesForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<Messages>;
        calendarForUser(userPrincipalName: string, callback: (events: CalendarEvent, error: Error) => void, odataQuery?: string): void;
        calendarForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<CalendarEvents>;
        memberOfForUser(userPrincipalName: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string): void;
        memberOfForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<Messages>;
        managerForUser(userPrincipalName: string, callback: (manager: Kurve.User, error: Error) => void, odataQuery?: string): void;
        managerForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<User>;
        directReportsForUser(userPrincipalName: string, callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string): void;
        directReportsForUserAsync(userPrincipalName: string, odataQuery?: string): TypedPromise<Users>;
        profilePhotoForUser(userPrincipalName: string, callback: (photo: ProfilePhoto, error: Error) => void): void;
        profilePhotoForUserAsync(userPrincipalName: string): TypedPromise<ProfilePhoto>;
        profilePhotoValueForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void): void;
        profilePhotoValueForUserAsync(userPrincipalName: string): TypedPromise<any>;
        getAsync(url: string): TypedPromise<string>;
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
