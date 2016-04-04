/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    rb.                     actions, me, user, users
    rb.me                   actions, event, events, message, messages, calendarView
    rb.me.event             event(eventId:string) => Event
    rb.me.event("123")      action, attachment, attachments,

Each endpoint exposes the set of available actions, along with the necessary metadata for each action
    rb.me.event("123").actions
    -> GET, POST, PATCH, DELETE
    rb.me.event("123").actions.GET.pathWithQuery
    -> '/me/event/123/'
    (Soon this will include scope information as well)

Simply pass this metadata to the REST channel of your choice, e.g.
    MyRESTLibrary.get(rb.me.event("123").actions.pathWithQuery)

However if you provide an endpoint directly to our Graph implementation, it will infer the relevant types:
    graph.get(rb.me.event("123")).organizer.name

Certain endpoints have parameters that are encoded into the request either on the path or the querystring:
    rb.me.event("123")
    -> /me/events/123
    rb.me.calendarView([startDate],[endDate])
    -> /me/calendarView?startDate=[startDate]&endDate=[endDate]

You can add ODATA queries to these or any other endpoint:
    rb.me.messages("123").addQuery("$select=id,subject")
    -> /me/messages/123?$select=id,subject
    rb.me.calendarView([startDate],[endDate]).addQuery("$select=organizer")
    -> /me/calendarView?startDate=[startDate]&endDate=[endDate]&$select=organizer

The API mirrors the structure of the Graph paths:
    to access:
        /users/billba@microsoft.com/messages/123-456-789/attachments
    we do:
        rb.user("billba@microsoft.com").message("1234").attachments("6789")
    compare to the old Kurve-y way of doing things:
        graph.messageAttachmentForUser("billba@microsoft.com", "12345", "6789")

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/

import { UserDataModel, AttachmentDataModel, MessageDataModel, EventDataModel } from './models';

namespace RequestBuilder {

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

    function queryUnion(query1?:string, query2?:string) {
        if (query1)
            return query1 + (query2 ? "&" + query2 : "");
        else
            return query2;
    }

    var pathWithQuery = (path:string, query?:string) => path + (query ? "?" + query : "");

    export class Endpoint<Model> {
        public actions: Actions;
        private model:Model;        // need to reference Model to make type inference work
        constructor(protected path:string, protected query?:string) {
        }
        addQuery = (query?:string) => new AddQuery<Model>(this.path, queryUnion(this.query, query), this.actions);
        pathWithQuery = pathWithQuery(this.path, this.query);
    }

    export class AddQuery<Model> extends Endpoint<Model> {
        constructor(path:string, query:string, actions:Actions) {
            super(path, query);
            if (actions) {
                for (var verb in actions)
                    actions[verb].pathWithQuery = this.pathWithQuery;
                this.actions = actions;
            }
        }
    }

    export class Attachment extends Endpoint<AttachmentDataModel> {
        constructor(path:string, attachmentId:string) {
            super(path + "/attachments/" + attachmentId);
        }
        actions = {
            GET: new Action(this.path)
        };
    }

    var attachment = (path:string) => (attachmentId:string) => new Attachment(path, attachmentId);

    export class Attachments extends Endpoint<AttachmentDataModel> {
        constructor(path:string) {
            super(path + "/attachments");
        }
        actions = {
            GETCOLLECTION: new Action(this.path),
        };
    }

    var attachments = (path:string) => new Attachments(path);

    export class Message extends Endpoint<MessageDataModel> {
        constructor(path:string, messageId:string) {
            super(path + "/messages/" + messageId);
        }
        actions = {
            GET: new Action(this.path),
            POST: new Action(this.path)
        };
        attachment = attachment(this.path);
        attachments = attachments(this.path);
    }

    var message = (path:string) => (messageId:string) => new Message(path, messageId);

    export class Messages extends Endpoint<MessageDataModel> {
        constructor(path:string) {
            super(path + "/messages/");
        }
        actions = {
            GETCOLLECTION: new Action(this.path),
        };
    }

    var messages = (path:string) => new Messages(path);

    export class Event extends Endpoint<EventDataModel> {
        constructor(path:string, eventId:string) {
            super(path + "/events/");
        }
        actions = {
            GET: new Action(this.path),
        };
        attachment = attachment(this.path);
        attachments = attachments(this.path);
    }

    var event = (path:string) => (eventId:string) => new Event(path, eventId);

    export class Events extends Endpoint<EventDataModel> {
        constructor(path:string) {
            super(path + "/events/");
        }
        actions = {
            GETCOLLECTION: new Action(this.path),
        };
    }

    var events = (path:string) => new Events(path);

    export class CalendarView extends Endpoint<EventDataModel> {
        constructor(path:string, startDate:Date, endDate:Date) {
            super(path + "/calendarView", "startDateTime=" + startDate.toString() + "&endDateTime=" + endDate.toString()); // REVIEW need to restore toISOString()
        }
        actions = {
            GETCOLLECTION: new Action(pathWithQuery(this.path, this.query))
        }
    }

    var calendarView = (path:string) => (startDate:Date, endDate:Date) => new CalendarView(path, startDate, endDate);

    export class User extends Endpoint<UserDataModel> {
        constructor(path:string = "", userId?:string) {
            super(userId? path + "/users/" + userId : path + "/me");
        }
        actions = {
            GET: new Action(this.path),
        };
        message = message(this.path);
        messages = messages(this.path);
        event = event(this.path);
        events = events(this.path);
        calendarView = calendarView(this.path);
    }

    var me = new User();
    var user = (userId:string) => new User("", userId);

    export class Users extends Endpoint<UserDataModel> {
        constructor(path:string = "") {
            super(path + "/users");
        }
        actions = {
            GETCOLLECTION: new Action(this.path),
        };
    }

    var users = new Users();

    export class Root {
        me = me;
        user = user;
        users = users;
    }
}

interface Collection<Model> {
    collection:Model[];
    //  nextLink callback will go here
}

class MockGraph {
    getUntypedCollection(path:string, scopes:string[]):any {
        return {};
    }

    getTypedCollection<Model>(path:string, scopes:string[]):Collection<Model> {
        return this.getUntyped(path, scopes);
    }

    getCollection<Model>(endpoint:RequestBuilder.Endpoint<Model>):Collection<Model> {
        var action = endpoint.actions && endpoint.actions.GETCOLLECTION;
        if (!action) {
            console.log("no GETCOLLECTION endpoint, sorry!");
        } else {
            console.log("GETCOLLECTION path", action.pathWithQuery);
            return this.getTypedCollection<Model>(action.pathWithQuery, action.scopes);
        }
    }

    getUntyped(path:string, scopes:string[]):any {
        return {};
    }

    getTyped<Model>(path:string, scopes:string[]):Model {
        return;
    }

    get<Model>(endpoint:RequestBuilder.Endpoint<Model>):Model {
        var action = endpoint.actions && endpoint.actions.GET;
        if (!action) {
            console.log("no GET endpoint, sorry!");
        } else {
            console.log("GET path", action.pathWithQuery);
            return this.getTyped<Model>(action.pathWithQuery, action.scopes);
        }
    }

    postUntyped(path:string, scopes:string[], request:any):any {
        return {};
    }

    postTyped<Model>(path:string, scopes:string[], request:Model):Model {
        return this.postUntyped(path, scopes, request);
    }

    post<Model>(endpoint:RequestBuilder.Endpoint<Model>, request:Model):Model {
        var action = endpoint.actions && endpoint.actions.POST;
        if (!action) {
            console.log("no POST endpoint, sorry!");
        } else {
            console.log("POST path", action.pathWithQuery);
            return this.postTyped<Model>(action.pathWithQuery, action.scopes, request);
        }
    }

    patchUntyped(path:string, scopes:string[], request:any):any {
        return {};
    }

    patchTyped<Model>(path:string, scopes:string[], request:Model):Model {
        return this.patchUntyped(path, scopes, request);
    }

    patch<Model>(endpoint:RequestBuilder.Endpoint<Model>, request:Model):Model {
        var action = endpoint.actions && endpoint.actions.PATCH;
        if (!action) {
            console.log("no PATCH endpoint, sorry!");
        } else {
            console.log("PATCH path", action.pathWithQuery);
            return this.patchTyped<Model>(action.pathWithQuery, action.scopes, request);
        }
    }

    deleteUntyped(path:string, scopes:string[]):void {
    }

    deleteTyped<Model>(path:string, scopes:string[]):void {
    }

    delete<Model>(endpoint:RequestBuilder.Endpoint<Model>):void {
        var action = endpoint.actions.DELETE;
        if (!action) {
            console.log("no DELETE endpoint, sorry!");
        } else {
            console.log("DELETE path", action.pathWithQuery);
            this.getTyped<Model>(action.pathWithQuery, action.scopes);
        }
    }

}

var rb = new RequestBuilder.Root();
var graph = new MockGraph();
