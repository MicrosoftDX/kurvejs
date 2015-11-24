declare module Sample {
    class App {
        private tenant;
        private clientId;
        private redirectUri;
        private identity;
        private graph;
        constructor();
        private logout();
        private loadUsersWithPaging();
        private loadUsersWithOdataQuery(query);
        private loadUserMe();
        private loadUserMessages();
        private isLoggedIn();
        private whoAmI();
        private getUsersCallback(users, error);
        private messagesCallback(messages, error);
    }
}

declare module Kurve {
    class Graph {
        private req;
        private state;
        private nonce;
        private accessToken;
        private tenantId;
        private KurveIdentity;
        private defaultResourceID;
        private baseUrl;
        constructor(tenantId: string, identityInfo: {
            identity: Identity;
        });
        constructor(tenantId: string, identityInfo: {
            defaultAccessToken: string;
        });
        meAsync(odataQuery?: string): Promise;
        me(callback: (user: any, error: string) => void, odataQuery?: string): void;
        usersAsync(odataQuery?: string): Promise;
        users(callback: (users: any, error: string) => void, odataQuery?: string): void;
        getAsync(url: string): Promise;
        get(url: string, callback: (response: string) => void): void;
        private getUsers(urlString, callback);
        private getUser(urlString, callback);
        private addAccessTokenAndSend(xhr);
        private decorateUserObject(user);
        private decorateMessageObject(message);
        private getMessages(urlString, callback, odataQuery?);
        private buildMeUrl();
        private buildUsersUrl();
    }
}

declare module Kurve {
    class Identity {
        authContext: any;
        config: any;
        isCallback: boolean;
        clientId: string;
        tenantId: string;
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
        constructor(tenantId?: string, clientId?: string, redirectUri?: string);
        private decodeIdToken(idToken);
        private decodeAccessToken(accessToken, resource);
        getIdToken(): any;
        isLoggedIn(): boolean;
        private renewIdToken();
        getAccessTokenAsync(resource: string): Promise;
        getAccessToken(resource: string, callback: (token: string, error: string) => void): void;
        loginAsync(): Promise;
        login(callback: (error: string) => void): void;
        logOut(): void;
        private base64Decode(encodedString);
        private generateNonce();
    }
}

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
