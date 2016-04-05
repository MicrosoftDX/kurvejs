/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    graph.                     me, user, users
    graph.me                   event, events, message, messages, calendarView, GET, PATCH, DELETE
    graph.me.event             event(eventId:string) => Event
    graph.me.event("123")      attachment, attachments, GET, PATCH, DELETE

Each endpoint exposes the set of available REST verbs through strongly typed methods:
    graph.me.GET():UserDataModel
    graph.me.events.GETCOLLECTION():EventDataModel
    graph.me.events.POST(event:EventDataModel):EventDataModel

Certain endpoints have parameters that are encoded into the request either in the path or the querystring:
    graph.me.event("123")
    -> /me/events/123
    graph.me.calendarView([startDate],[endDate])
    -> /me/calendarView?startDate=[startDate]&endDate=[endDate]

You can add ODATA queries through the REST methods:
    graph.me.messages("123").GET("$select=id,subject")
    -> GET "/me/messages/123?$select=id,subject""
    graph.me.calendarView([startDate],[endDate]).GETCOLLECTION("$select=organizer")
    -> GETCOLLECTION "/me/calendarView?startDate=[startDate]&endDate=[endDate]&$select=organizer""

The API mirrors the structure of the Graph paths:
    to access:
        /users/billba@microsoft.com/messages/123-456-789/attachments
    we do:
        graph.user("billba@microsoft.com").message("1234").attachments("6789")
    compare to the old Kurve-y way of doing things:
        graph.messageAttachmentForUser("billba@microsoft.com", "12345", "6789")

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/

import { Promise } from "./promises";
import { Graph } from "./graph";
import { Error } from "./identity";
import { UserDataModel, AttachmentDataModel, MessageDataModel, EventDataModel, MailFolderDataModel } from './models';

export interface Collection<Model> {
    objects:Model[];
    nextLink?:any;
    //  nextLink callback will go here
}

export var pathWithQuery = (path:string, query1?:string, query2?:string) => {
    var query = (query1 ? query1 + (query2 ? "&" + query2 : "" ) : query2); 
    return path + (query ? "?" + query : "");
}

export abstract class Endpoint {
    constructor(protected graph:Graph, protected path:string, protected query?:string) {
    }
//  pathWithQuery = (query?:string) => pathWithQuery(this.path, this.query, query);
}

export class Attachment extends Endpoint {
    constructor(protected graph:Graph, path:string, attachmentId:string) {
        super(graph, path + "/attachments/" + attachmentId);
    }

    GET = this.graph.GET<AttachmentDataModel>(this.path, this.query);
/*    
    PATCH = this.graph.PATCH<AttachmentDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<AttachmentDataModel>(this.path, this.query);
*/
}

var attachment = (graph:Graph, path:string) => (attachmentId:string) => new Attachment(graph, path, attachmentId);

export class Attachments extends Endpoint {
    constructor(protected graph:Graph, path:string) {
        super(graph, path + "/attachments");
    }

    GETCOLLECTION = this.graph.GETCOLLECTION<AttachmentDataModel>(this.path, this.query);
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

var attachments = (graph:Graph, path:string) => new Attachments(graph, path);

export class Message extends Endpoint {
    constructor(protected graph:Graph, path:string, messageId:string) {
        super(graph, path + "/messages/" + messageId);
    }

    GET = this.graph.GET<MessageDataModel>(this.path, this.query);
/*
    PATCH = this.graph.PATCH<MessageDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<MessageDataModel>(this.path, this.query);
*/
    attachment = attachment(this.graph, this.path);
    attachments = attachments(this.graph, this.path);
}

var message = (graph:Graph, path:string) => (messageId:string) => new Message(graph, path, messageId);

export class Messages extends Endpoint {
    constructor(protected graph:Graph, path:string) {
        super(graph, path + "/messages/");
    }

    GETCOLLECTION = this.graph.GETCOLLECTION<MessageDataModel>(this.path, this.query);
/*
    POST = this.graph.POST<MessageDataModel>(this.path, this.query);
*/
}

var messages = (graph:Graph, path:string) => new Messages(graph, path);

export class Event extends Endpoint {
    constructor(protected graph:Graph, path:string, eventId:string) {
        super(graph, path + "/events/");
    }

    GET = this.graph.GET<EventDataModel>(this.path, this.query);
/*
    PATCH = this.graph.PATCH<EventDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<EventDataModel>(this.path, this.query);
*/
    attachment = attachment(this.graph, this.path);
    attachments = attachments(this.graph, this.path);
}

var event = (graph:Graph, path:string) => (eventId:string) => new Event(graph, path, eventId);

export class Events extends Endpoint {
    constructor(protected graph:Graph, path:string) {
        super(graph, path + "/events/");
    }

    GETCOLLECTION = this.graph.GETCOLLECTION<EventDataModel>(this.path, this.query);
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

var events = (graph:Graph, path:string) => new Events(graph, path);

export class CalendarView extends Endpoint {
    constructor(protected graph:Graph, path:string, startDate:Date, endDate:Date) {
        super(graph, path + "/calendarView", "startDateTime=" + startDate.toISOString() + "&endDateTime=" + endDate.toISOString());
    }

    GETCOLLECTION = this.graph.GETCOLLECTION<EventDataModel>(this.path, this.query);
}

var calendarView = (graph:Graph, path:string) => (startDate:Date, endDate:Date) => new CalendarView(graph, path, startDate, endDate);

export class MailFolders extends Endpoint {
    constructor(protected graph:Graph, path:string) {
        super(graph, path + "/mailFolders");
    }

    GETCOLLECTION = this.graph.GETCOLLECTION<MailFolderDataModel>(this.path, this.query);
}

export class User extends Endpoint {
    constructor(protected graph:Graph, path:string = "", userId?:string) {
        super(graph, userId? path + "/users/" + userId : path + "/me");
    }

    GET = this.graph.GET<UserDataModel>(this.path, this.query);
/*
    PATCH = this.graph.PATCH<UserDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<UserDataModel>(this.path, this.query);
*/
    message = message(this.graph, this.path);
    messages = messages(this.graph, this.path);
    event = event(this.graph, this.path);
    events = events(this.graph, this.path);
    calendarView = calendarView(this.graph, this.path);
    mailFolders = new MailFolders(this.graph, this.path)
}

export class Users extends Endpoint {
    constructor(protected graph:Graph, path:string = "") {
        super(graph, path + "/users");
    }

    GETCOLLECTION = this.graph.GETCOLLECTION<UserDataModel>(this.path, this.query);
/*
    POST = this.graph.POST<UserDataModel>(this.path, this.query);
*/
}

/*
export class Graph {
    GetCollection<Model>(path:string, scopes?:string[]):Collection<Model> {
        console.log("GETCOLLECTION", path);
        return {objects:[]};
    }


    Get<Model>(path:string, scopes?:string[]): Promise<Model, Error> {
        console.log("GET", path);
        return {} as Model;
    }

    Post<Model>(object:Model, path:string, scopes?:string[]):Model {
        console.log("POST", path);
        return {} as Model;
    }

    Patch<Model>(object:Model, path:string, scopes?:string[]):Model {
        console.log("PATCH", path);
        return {} as Model;
    }

    Delete<Model>(path:string, scopes?:string[]):void {
        console.log("DELETE     ", path);
    }
    
    POST = <Model>(path:string, queryT?:string, scopes?:string[]) => (object:Model, query?:string) => this.Post<Model>(object, pathWithQuery(path, queryT, query), scopes);
    PATCH = <Model>(path:string, queryT?:string, scopes?:string[]) => (object:Model, query?:string) => this.Patch<Model>(object, pathWithQuery(path, queryT, query), scopes);
    DELETE = <Model>(path:string, queryT?:string, scopes?:string[]) => (query?:string) => this.Delete<Model>(pathWithQuery(path, queryT, query), scopes);
}
*/

