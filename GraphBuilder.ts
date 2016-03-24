
/*

QueryBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    root.                      me, path, users, endpoints
    root.me                    event, events, message, messages, path
    root.me.event              event(eventId:string) => Event
    
The API mirrors the structure of the Graph paths:
    to access:
        /users/billba@microsoft.com/messages/123-456-789/attachments
    instead of the Kurve-y way of doing things:
        graph.messageAttachmentsForUser("billba@microsoft.com", "123-456-789")
    we do:
        root.user("billba@microsoft.com").message("123-456-789").attachments

Different nodes surface appropriate functionality:
    root.user("bill").         delete, event, events, get, message, messages, path
    root.users.                get, path

Each node surfaces a path which can be passed to the channel of your choice to access the MS Graph: 
    root.user("billba@microsoft.com").message("123-456-789").attachments.pathWithQuery("$select=id,inline")
    => '/users/billba@microsoft.com/messages/123-456-789/attachments?$select=id,inline'

Or use the convenient built-in strongly-typed REST methods!
    root.user("bill").messages().get()        => Graph.MessageDataModel[]
    root.user("bill").event("123").get()      => Graph.EventDataModel

Each REST method has available to it the appropriate path via "this.pathWithQuery".
A similar approach would be used to convey scopes down the call chain.

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/


module QueryBuilder {
        
    export class Endpoint {
        constructor(public path: string, public scopes: string[]){}
    }

    export class Endpoints<T> {
        get: Endpoint;
        post: Endpoint;
        patch: Endpoint;
        delete: Endpoint;

        constructor(path:string, get_scopes?: string[], post_scopes?: string[], patch_scopes?: string[], delete_scopes?: string[]) {
            if (get_scopes)
                this.get = new Endpoint(path, get_scopes);
            if (post_scopes)
                this.post = new Endpoint(path, post_scopes);
            if (patch_scopes)
                this.patch = new Endpoint(path, patch_scopes);
            if (delete_scopes)
                this.delete = new Endpoint(path, delete_scopes);
        }
    }
    
    export class Get<Model> {
        constructor (public path: string, type?: Model) {}
    }

    interface IGet<Model> {
        (query?:string): Get<Model>;
    }
    
    function get<Model>(query?:string) {
        return new Get<Model>(this.pathWithQuery);
    }
    
    export abstract class QueryNode<T> {
        constructor(protected path: string = "", protected query?: string) {}
        get pathWithQuery() { return this.path + (this.query ? "?" + this.query : ""); }
        endpoints:Endpoints<T>;
    }
    
    export class Attachment extends QueryNode<Kurve.AttachmentDataModel> {
        endpoints = new Endpoints<Kurve.AttachmentDataModel>(this.path);
    }

    export class Attachments extends QueryNode<Kurve.AttachmentDataModel[]> {
        endpoints = new Endpoints<Kurve.AttachmentDataModel[]>(this.path);
    }
    
    abstract class HasAttachments<T> extends QueryNode<T> {
        attachment = (attachmentId:string, query?:string) => new Attachment(this.path + "/attachments/" + attachmentId, query);
        attachments = (query?:string) => new Attachments(this.path + "/attachments", query);   
    }

    export class Message extends QueryNode<Kurve.MessageDataModel> {
        endpoints = new Endpoints<Kurve.MessageDataModel>(this.path);
    }

    export class Messages extends QueryNode<Kurve.MessageDataModel[]> {
        endpoints = new Endpoints<Kurve.MessageDataModel[]>(this.path);
    }
    
    export class Event extends QueryNode<Kurve.EventDataModel> {
        attachment = (attachmentId:string, query?:string) => new Attachment(this.path + "/attachments/" + attachmentId, query);
        attachments = (query?:string) => new Attachments(this.path + "/attachments", query);   
        endpoints = new Endpoints<Kurve.EventDataModel>(this.path);
        }

    export class Events extends QueryNode<Kurve.EventDataModel[]> {
        endpoints = new Endpoints<Kurve.EventDataModel[]>(this.path);
    }

    export class User extends QueryNode<Kurve.UserDataModel> {
        public message = (messageId: string) => new Message(this.path + "/messages/" + messageId);
        public messages = new Messages(this.path + "/messages");
        public event = (eventId: string) => new Event(this.path + "/events/" + eventId);
        public events = new Events(this.path + "/events");
        public calendarView = (startDate:Date, endDate:Date) => new Events(this.path + "/calendarView", ""); // REVIEW incorporate start & end dates
        endpoints = new Endpoints<Kurve.UserDataModel>(this.path);
    }

    export class Users extends QueryNode<Kurve.UserDataModel[]> {
        endpoints = new Endpoints<Kurve.UserDataModel[]>(this.path);
    }

    export class Root {
        me = new User("/me");
        user = (userId:string) => new User("/users/" + userId);
        users = new Users("/users/");
    }

}


class MockGraph {
    get<T>(query:QueryBuilder.QueryNode<T>):T {
        if (!query.endpoints.get) {
            console.log("no GET endpoint, sorry!");
            return;
        }
        console.log("path", query.endpoints.get.path);
        return {} as T;
    }    
    post<T>(query:QueryBuilder.QueryNode<T>, request:T):void {
        if (!query.endpoints.post) {
            console.log("no POST endpoint, sorry!");
            return;
        }
        console.log("path", query.endpoints.get.path);
    }    
}

var root = new QueryBuilder.Root();
var graph = new MockGraph();

graph.get(root.me.message("123")).body.content
graph.post(root.me.message("123"), new Kurve.MessageDataModel());
