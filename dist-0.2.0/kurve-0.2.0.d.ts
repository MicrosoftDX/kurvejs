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
        private deferred;
        constructor(deferred: Deferred);
        then(doneFilter: Function, failFilter?: Function, progressFilter?: Function): Promise;
        done(...callbacks: Function[]): Promise;
        fail(...callbacks: Function[]): Promise;
        always(...callbacks: Function[]): Promise;
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
        meAsync(odataQuery?: string): Promise;
        me(callback: (user: any, error: Error) => void, odataQuery?: string): void;
        userAsync(userId: string): Promise;
        user(userId: string, callback: (users: any, error: Error) => void): void;
        usersAsync(odataQuery?: string): Promise;
        users(callback: (users: any, error: Error) => void, odataQuery?: string): void;
        groupAsync(groupId: string): Promise;
        group(groupId: string, callback: (group: any, error: Error) => void): void;
        groups(callback: (groups: any, error: Error) => void, odataQuery?: string): void;
        groupsAsync(odataQuery?: string): Promise;
        getAsync(url: string): Promise;
        get(url: string, callback: (response: string, error: Error) => void, responseType?: string): void;
        private generateError(xhr);
        private getUsers(urlString, callback);
        private getUser(urlString, callback);
        private addAccessTokenAndSend(xhr, callback);
        private decorateUserObject(user);
        private decorateMessageObject(message);
        private decorateGroupObject(message);
        private decoratePhotoObject(message);
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
