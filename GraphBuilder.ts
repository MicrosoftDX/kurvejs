/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

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


module RequestBuilder {

    export class Endpoint {
        constructor(public path: string, public scopes: string[]){}
    }

    var Verbs = ["GET", "POST", "PATCH", "DELETE"];

    export class Endpoints<T> {
        constructor(path:string, query?:string, getScopes?: string[], postScopes?: string[], patchScopes?: string[], deleteScopes?: string[]) {
            if (query)
                path = path + "?" + query;
            if (getScopes)
                this["GET"] = new Endpoint(path, getScopes);
            if (postScopes)
                this["POST"] = new Endpoint(path, postScopes);
            if (patchScopes)
                this["PATCH"] = new Endpoint(path, patchScopes);
            if (deleteScopes)
                this["DELETE"] = new Endpoint(path, deleteScopes);
        }
    }

    export abstract class Node<T> {
        constructor(protected path: string = "", protected query?: string, public endpoints?: Endpoints<T>) {
        }
    }

    function queryUnion(query1?:string, query2?:string) {
        if (query1)
            return query1 + (query2 ? "&" + query2 : "");
        else
            return query2; 
    }

    export class AddQuery<Model> extends Node<Model> {
        constructor(protected path: string = "", protected query?: string, public endpoints?: Endpoints<Model>) {
            super(path, query, endpoints);
            var path = this.path + (this.query ? "?" + this.query : "");
            for (var verb in Verbs)
                if (endpoints[verb])
                    endpoints[verb] = path;
        }
    }

    abstract class NodeWithQuery<Model> extends Node<Model> {
        addQuery = (query?:string) => new AddQuery<Model>(this.path, queryUnion(this.query, query), this.endpoints);
    }

    export class Attachment extends Node<Kurve.AttachmentDataModel> {
        endpoints = new Endpoints<Kurve.AttachmentDataModel>(this.path, this.query);
    }

    export class Attachments extends Node<Kurve.AttachmentDataModel[]> {
        endpoints = new Endpoints<Kurve.AttachmentDataModel[]>(this.path);
    }
    
    abstract class NodeWithAttachments<T> extends NodeWithQuery<T> {
        attachment = (attachmentId:string) => new Attachment(this.path + "/attachments/" + attachmentId);
        attachments = new Attachments(this.path + "/attachments");
    }

    export class Message extends NodeWithAttachments<Kurve.MessageDataModel> {
        endpoints = new Endpoints<Kurve.MessageDataModel>(this.path);
    }

    export class Messages extends NodeWithQuery<Kurve.MessageDataModel[]> {
        endpoints = new Endpoints<Kurve.MessageDataModel[]>(this.path);
    }
    
    export class Event extends NodeWithAttachments<Kurve.EventDataModel> {
        endpoints = new Endpoints<Kurve.EventDataModel>(this.path);
    }

    export class Events extends NodeWithQuery<Kurve.EventDataModel[]> {
        endpoints = new Endpoints<Kurve.EventDataModel[]>(this.path);
    }

    export class User extends NodeWithQuery<Kurve.UserDataModel> {
        message = (messageId: string) => new Message(this.path + "/messages/" + messageId);
        messages = new Messages(this.path + "/messages");
        event = (eventId: string) => new Event(this.path + "/events/" + eventId);
        events = new Events(this.path + "/events");
        calendarView = (startDate:Date, endDate:Date) => new Events(this.path + "/calendarView", ""); // REVIEW incorporate start & end dates
        endpoints = new Endpoints<Kurve.UserDataModel>(this.path);
    }

    export class Users extends NodeWithQuery<Kurve.UserDataModel[]> {
        endpoints = new Endpoints<Kurve.UserDataModel[]>(this.path);
    }

    export class Root {
        me = new User("/me");
        user = (userId:string) => new User("/users/" + userId);
        users = new Users("/users/");
    }
}

class MockGraph {
    get<T>(query:RequestBuilder.Node<T>):T {
        var endpoint = query.endpoints["GET"];
        if (!endpoint) {
            console.log("no GET endpoint, sorry!");
            return;
        }
        console.log("GET path", endpoint.path);
        return {} as T;
    }    
    post<T>(query:RequestBuilder.Node<T>, request:T):void {
        var endpoint = query.endpoints["POST"];
        if (!endpoint) {
            console.log("no POST endpoint, sorry!");
            return;
        }
        console.log("POST path", endpoint.path);
    }    
}

var root = new RequestBuilder.Root();
var graph = new MockGraph();

//root.me().event("123").endpoints.get
//graph.get(root.me().message("123")).body.content
//graph.post(root.me().message("123"), new Kurve.MessageDataModel());
