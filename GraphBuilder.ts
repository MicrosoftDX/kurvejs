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
    root.user("bill").messages().get()        => Graph.MessageDataModel
    root.user("bill").event("123").get()      => Graph.EventDataModel

Each REST method has available to it the appropriate path via "this.pathWithQuery".
A similar approach would be used to convey scopes down the call chain.

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/

// A mock of Kurve for testing purposes.

module Kurve {
    export class UserDataModel {
        id:string;
        userField = "I am a user.";
    }

    export class MessageDataModel {
        id:string;
        messageField = "I am a message.";
    }
    
    export class EventDataModel {
        id:string;
        eventField = "I am an event.";
    }
    
    export class AttachmentDataModel {
        id:string;
        attachmentField = "I am an attachment.";
    }
}

module RequestBuilder {

    export class Action {
        constructor(public pathWithQuery:string, public scopes?:string[]){}
    }
    
    interface Actions<Model> {
        GET?:Action;
        GETCOLLECTION?:Action;
        POST?:Action;
        DELETE?:Action;
        PATCH?:Action;
    }

    export abstract class Node<Model> {
        actions:Actions<Model>;
        constructor(protected path: string = "", protected query?: string, actions?: Actions<Model>) {
        }
        protected get pathWithQuery() { return this.path + (this.query ? "?" + this.query : "") }
    }

    function queryUnion(query1?:string, query2?:string) {
        if (query1)
            return query1 + (query2 ? "&" + query2 : "");
        else
            return query2; 
    }

    export class AddQuery<Model> extends Node<Model> {
        constructor(protected path: string = "", protected query?: string, actions?: Actions<Model>) {
            super(path, query, actions);
            if (actions)
                for (var verb in actions)
                    actions[verb].pathWithQuery = this.pathWithQuery;
        }
    }

    abstract class NodeWithQuery<Model> extends Node<Model> {
        addQuery = (query?:string) => new AddQuery<Model>(this.path, queryUnion(this.query, query), this.actions);
    }

    export class Attachment extends Node<Kurve.AttachmentDataModel> {
    }

    export class Attachments extends Node<Kurve.AttachmentDataModel> {
    }
    
    export abstract class NodeWithAttachments<T> extends NodeWithQuery<T> {
        attachment = (attachmentId:string) => new Attachment(this.path + "/attachments/" + attachmentId);
        attachments = new Attachments(this.path + "/attachments");
    }

    export class Message extends NodeWithAttachments<Kurve.MessageDataModel> {
        actions = {
            GET: new Action(this.pathWithQuery),
            POST: new Action(this.pathWithQuery)
        }
    }

    export class Messages extends NodeWithQuery<Kurve.MessageDataModel> {
        actions = {
            GETCOLLECTION: new Action(this.pathWithQuery),
            POST: new Action(this.pathWithQuery)
        }
    }
    
    export class Event extends NodeWithAttachments<Kurve.EventDataModel> {
    }

    export class Events extends NodeWithQuery<Kurve.EventDataModel[]> {
    }

    export class User extends NodeWithQuery<Kurve.UserDataModel> {
        message = (messageId: string) => new Message(this.path + "/messages/" + messageId);
        messages = new Messages(this.path + "/messages");
        event = (eventId: string) => new Event(this.path + "/events/" + eventId);
        events = new Events(this.path + "/events");
        calendarView = (startDate:Date, endDate:Date) => new Events(this.path + "/calendarView", "startDateTime=" + startDate.toISOString() + "&endDateTime=" + endDate.toISOString());
    }

    export class Users extends NodeWithQuery<Kurve.UserDataModel> {
    }

    export class Root {
        me = new User("/me");
        user = (userId:string) => new User("/users/" + userId);
        users = new Users("/users/");
    }
}

interface Collection<Model> {
    collection:Model[];
    //  nextLink callback will go here 
}

class MockGraph {
    getCollection<Node extends RequestBuilder.Node<Model>, Model>(endpoint:Node):Collection<Model> {
        var action = endpoint.actions.GETCOLLECTION;
        if (!action) {
            console.log("no GETCOLLECTION endpoint, sorry!");
        } else {
            console.log("GETCOLLECTION path", action.pathWithQuery);
            return {} as Collection<Model>;
        }
    }

    get<Node extends RequestBuilder.Node<Model>, Model>(endpoint:RequestBuilder.Node<Model>):Model {
        var action = endpoint.actions.GET;
        if (!action) {
            console.log("no GET endpoint, sorry!");
        } else {
            console.log("GET path", action.pathWithQuery);
            return {} as Model;
        }
    }
    
    post<Node extends RequestBuilder.Node<Model>, Model>(endpoint:RequestBuilder.Node<Model>, request:Model):void {
        var action = endpoint.actions.POST;
        if (!action) {
            console.log("no POST endpoint, sorry!");
        } else {
            console.log("POST path", action.pathWithQuery);
        }
    }

    patch<Node extends RequestBuilder.Node<Model>, Model>(endpoint:RequestBuilder.Node<Model>, request:Model):void {
        var action = endpoint.actions.PATCH;
        if (!action) {
            console.log("no PATCH endpoint, sorry!");
        } else {
            console.log("PATCH path", action.pathWithQuery);
        }
    }
    
    delete<Node extends RequestBuilder.Node<Model>, Model>(endpoint:RequestBuilder.Node<Model>, request:Model):void {
        var action = endpoint.actions.DELETE;
        if (!action) {
            console.log("no DELETE endpoint, sorry!");
        } else {
            console.log("DELETE path", action.pathWithQuery);
        }
    }

}

var rb = new RequestBuilder.Root();
var graph = new MockGraph();

graph.get(rb.me.message("123"))

//root.me().event("123").endpoints.get
//graph.get(root.me().message("123")).body.content
//graph.post(root.me().message("123"), new Kurve.MessageDataModel());
