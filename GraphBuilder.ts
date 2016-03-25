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
    
    interface Actions {
        GET?:Action;
        GETCOLLECTION?:Action;
        POST?:Action;
        DELETE?:Action;
        PATCH?:Action;
    }

    export abstract class Node<Model> {
        private model:Model; // we need to reference Model somewhere to make type inference work 
        actions:Actions;
        constructor(protected path:string = "", protected query?:string) {
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
        constructor(protected path:string, protected query?:string, actions?:Actions) {
            super(path, query);
            if (actions) {
                this.actions = actions;
                for (var verb in actions)
                    actions[verb].pathWithQuery = this.pathWithQuery;
            }
        }
    }

    export abstract class NodeWithQuery<Model> extends Node<Model> {
        addQuery = (query?:string) => new AddQuery<Model>(this.path, queryUnion(this.query, query), this.actions);
    }

    export class Attachment<Model> extends NodeWithQuery<Model> {
    }

    export class Attachments<Model> extends NodeWithQuery<Model> {
    }
    
    export abstract class NodeWithAttachments<Model> extends NodeWithQuery<Model> {
        attachment = (attachmentId:string) => new Attachment<Kurve.AttachmentDataModel>(this.path + "/attachments/" + attachmentId);
        attachments = new Attachments<Kurve.AttachmentDataModel>(this.path + "/attachments");
    }

    export class Message<Model> extends NodeWithAttachments<Model> {
        actions:Actions = {
            GET: new Action(this.pathWithQuery),
            POST: new Action(this.pathWithQuery)
        }
    }

    export class Messages<Model> extends NodeWithQuery<Model> {
        actions:Actions = {
            GETCOLLECTION: new Action(this.pathWithQuery),
            POST: new Action(this.pathWithQuery)
        }
    }
    
    export class Event<Model> extends NodeWithAttachments<Model> {
    }

    export class Events<Model> extends NodeWithQuery<Model> {
    }

    export class User<Model> extends NodeWithQuery<Model> {
        message = (messageId:string) => new Message<Kurve.MessageDataModel>(this.path + "/messages/" + messageId);
        messages = new Messages<Kurve.MessageDataModel>(this.path + "/messages");
        event = (eventId:string) => new Event<Kurve.EventDataModel>(this.path + "/events/" + eventId);
        events = new Events<Kurve.EventDataModel>(this.path + "/events");
        calendarView = (startDate:Date, endDate:Date) => new Events<Kurve.EventDataModel>(this.path + "/calendarView", "startDateTime=" + startDate.toISOString() + "&endDateTime=" + endDate.toISOString());
    }

    export class Users<Model> extends NodeWithQuery<Model> {
    }

    export class Root {
        me = new User<Kurve.UserDataModel>("/me");
        user = (userId:string) => new User<Kurve.UserDataModel>("/users/" + userId);
        users = new Users<Kurve.UserDataModel>("/users/");
    }
}

interface Collection<Model> {
    collection:Model[];
    //  nextLink callback will go here 
}

class MockGraph {
    getCollection<Model>(endpoint:RequestBuilder.Node<Model>):Collection<Model> {
        var action = endpoint.actions && endpoint.actions.GETCOLLECTION;
        if (!action) {
            console.log("no GETCOLLECTION endpoint, sorry!");
        } else {
            console.log("GETCOLLECTION path", action.pathWithQuery);
            return {collection:[]} as Collection<Model>;
        }
    }

    get<Model>(endpoint:RequestBuilder.Node<Model>):Model {
        var action = endpoint.actions && endpoint.actions.GET;
        if (!action) {
            console.log("no GET endpoint, sorry!");
        } else {
            console.log("GET path", action.pathWithQuery);
            return {} as Model;
        }
    }
    
    post<Model>(endpoint:RequestBuilder.Node<Model>, request:Model):void {
        var action = endpoint.actions && endpoint.actions.POST;
        if (!action) {
            console.log("no POST endpoint, sorry!");
        } else {
            console.log("POST path", action.pathWithQuery);
        }
    }

    patch<Model>(endpoint:RequestBuilder.Node<Model>, request:Model):void {
        var action = endpoint.actions && endpoint.actions.PATCH;
        if (!action) {
            console.log("no PATCH endpoint, sorry!");
        } else {
            console.log("PATCH path", action.pathWithQuery);
        }
    }
    
    delete<Model>(endpoint:RequestBuilder.Node<Model>):void {
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

/*
var foo = rb.me.message("123");
var bar = rb.me.messages.addQuery("foo")

graph.get(foo).messageField
graph.getCollection(bar);
graph.post(foo, {} as Kurve.MessageDataModel);
graph.delete(foo);
graph.patch(foo, {} as Kurve.MessageDataModel);
*/