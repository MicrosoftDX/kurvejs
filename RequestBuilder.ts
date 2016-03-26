/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    rb.                     actions, me, user, users
    rb.me                   actions, event, events, message, messages, calendarView
    rb.me.event             event(eventId:string) => Event
    rb.me.event("123")      actions, attachment, attachments
    
Each endpoint in the graph surfaces an "actions" object
    rb.me.event("123").actions

The "actions" object contains metadata for each supported REST verb, e.g.
    rb.me.event("123").actions.GET.pathWithQuery = '/me/event/123/'
    (Soon this will include scope information as well)

Simply pass this metadata to the REST channel of your choice, e.g.
    MyRESTLibrary.get(rb.me.event("123").actions.restWithQuery)

However if you provide an endpoint directly to our Graph implementation, it will infer the relevant types: 
    graph.get(rb.me.event("123")).organizer.name
    
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

// A mock of Kurve for testing purposes.

namespace Kurve {
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

    export abstract class Node<Model> {
        private model:Model;                        // need to reference Model to make type inference work
        constructor(public actions?:Actions) {}     // need this in the base class so that its available externally  
    }

    function queryUnion(query1?:string, query2?:string) {
        if (query1)
            return query1 + (query2 ? "&" + query2 : "");
        else
            return query2; 
    }

    var pathWithQuery = (path:string, query?:string) => path + (query ? "?" + query : "");

   // Three types of nodes:
   //   PossibleEndpoint        /me.messages                                    Endpoint without query. Might just be an intermediary node.
   //   EndpointWithQuery       /me.calendarView?$startDate=X&endDate=Y         Endpoint with built-in query
   //   AddQuery                /me.messages?$select=id,bodyPreview             Either of the above with added query
   //                           /me.calendarView?$startDate=X&endDate=Y&r=1

    export class AddQuery<Model> extends Node<Model> {
        constructor(path:string, public actions:Actions, query?:string) {
            super(actions);
            if (actions)
                for (var verb in actions)
                    actions[verb].pathWithQuery = pathWithQuery(path, query);
        }
    }

    export abstract class Endpoint<Model> extends Node<Model> {
        constructor(protected path:string, public actions:Actions, protected query?:string) {
            super(actions);
        }
        addQuery = (query?:string) => new AddQuery<Model>(this.path, this.actions, queryUnion(this.query, query));
        protected get pathWithQuery() {return pathWithQuery(this.path, this.query)}
    }
    
    export class Attachment<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GET: new Action(path),
            });
        }
    }
    
    var attachment = (path:string) => (attachmentId:string) => new Attachment<Kurve.AttachmentDataModel>(path + "/attachments/" + attachmentId);

    export class Attachments<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GETCOLLECTION: new Action(path),
            });
        }
    }

    var attachments = (path:string) => new Attachments<Kurve.AttachmentDataModel>(path + "/attachments");

    export class Message<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GET: new Action(path),
                POST: new Action(path)
            });
        }
        attachment = attachment(this.path);
        attachments = attachments(this.path);
    }

    var message = (path:string) => (messageId:string) => new Message<Kurve.MessageDataModel>(path + "/messages/" + messageId);

    export class Messages<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GETCOLLECTION: new Action(path),
            });
        }
    }

    var messages = (path:string) => new Messages<Kurve.MessageDataModel>(path + "/messages");
    
    export class Event<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GET: new Action(path),
            });
        }
        attachment = attachment(this.path);
        attachments = attachments(this.path);
    }
    
    var event = (path:string) => (eventId:string) => new Event<Kurve.EventDataModel>(path + "/events/" + eventId);

    export class Events<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GETCOLLECTION: new Action(path),
            });
        }
    }
    
    var events = (path:string) => new Events<Kurve.EventDataModel>(path + "/events");   

    export class CalendarView<Model> extends Endpoint<Model> {
        constructor(path:string, startDate:Date, endDate:Date) {
            var query = "startDateTime=" + startDate.toString() + "&endDateTime=" + endDate.toString();
//          REVIEW need to restore toISOString()
            super(path, {
                GETCOLLECTION: new Action(pathWithQuery(path, query))
            }, query);
        }
    }
    
    var calendarView = (path:string) => (startDate:Date, endDate:Date) => new CalendarView<Kurve.EventDataModel>(path + "/calendarView", startDate, endDate);

    export class User<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GET: new Action(path),
            });
        }
        message = message(this.path);
        messages = messages(this.path);
        event = event(this.path);
        events = events(this.path);
        calendarView = calendarView(this.path);
    }

    var me = new User<Kurve.UserDataModel>("/me");
    var user = (userId:string) => new User<Kurve.UserDataModel>("/users/" + userId);

    export class Users<Model> extends Endpoint<Model> {
        constructor(protected path:string) {
            super(path, {
                GETCOLLECTION: new Action(path),
            });
        }
    }

    var users = new Users<Kurve.UserDataModel>("/users/");

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
var foo = rb.me.message("123").actions
var bar = rb.me.calendarView(new Date(), new Date()).addQuery("123")

graph.get(foo).messageField
graph.getCollection(bar).collection
graph.post(foo, {} as Kurve.MessageDataModel);
graph.delete(foo);
graph.patch(foo, {} as Kurve.MessageDataModel);
*/