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

export class Response<Model, N extends Node> {
    constructor(public raw:any, public self:N) {
    }

    get object() {
        return this.raw as Model;
    }
}

export class Collection<Model, N extends Node> {
    constructor(public raw:any, public self:N, public next:N) {
        let nextLink = this.raw["@odata.nextLink"];
        if (nextLink) {
            this.next.pathWithQuery = nextLink; 
        } else {
            this.next = undefined;
        }
    }

    get objects() {
        return (this.raw.value ? this.raw.value : [this.raw]) as Model[];
    }
}

let queryUnion = (query1:string, query2:string) => (query1 ? query1 + (query2 ? "&" + query2 : "" ) : query2); 

let pathWithQuery = (path:string, query1?:string, query2?:string) => {
    let query = queryUnion(query1, query2); 
    return path + (query ? "?" + query : "");
}

export abstract class Node {
    constructor(protected graph:Graph, protected path:string, protected query?:string) {
    }

    get pathWithQuery() {
        return pathWithQuery(this.path, this.query);
    }

    set pathWithQuery(pathWithQuery:string) {
        let i = pathWithQuery.indexOf("?");
        if (i == -1) {
            this.path = pathWithQuery;
            this.query = undefined;    
        } else {
            this.path = pathWithQuery.substring(0, i);
            this.query = pathWithQuery.substring(i + 1);
        }
    }
    
    odata = (query:string) => {
        this.query = queryUnion(this.query, query);
        return this;
    }
    orderby = (...fields:string[]) => this.odata(`$orderby=${fields.join(",")}`);
    top = (items:Number) => this.odata(`$top=${items.toString()}`);
    skip = (items:Number) => this.odata(`$skip=${items.toString()}`);
    filter = (query:string) => this.odata(`$filter=${query}`);
    expand = (...fields:string[]) => this.odata(`$expand=${fields.join(",")}`);
    select = (...fields:string[]) => this.odata(`$select=${fields.join(",")}`);
}

export class Attachment extends Node {
    constructor(graph:Graph, path:string="", attachmentId:string) {
        super(graph, path + "/attachments/" + attachmentId);
    }

    GetAttachment = () => this.graph.Get<AttachmentDataModel, Attachment>(this.pathWithQuery, this);
/*    
    PATCH = this.graph.PATCH<AttachmentDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<AttachmentDataModel>(this.path, this.query);
*/
}

export function _attachments():Attachments;
export function _attachments(attachmentId:string):Attachment;
export function _attachments(arg?:any):any {
    if (arg)
        return new Attachment(this.graph, this.path, arg);
    else
        return new Attachments(this.graph, this.path);
}

export class Attachments extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/attachments");
    }

    GetAttachments = () => this.graph.GetCollection<AttachmentDataModel, Attachments>(this.pathWithQuery, this, new Attachments(this.graph));
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

export class Message extends Node {
    constructor(graph:Graph, path:string="", messageId:string) {
        super(graph, path + "/messages/" + messageId);
    }

    attachments = _attachments;

    GetMessage = () => this.graph.Get<MessageDataModel, Message>(this.pathWithQuery, this);
/*
    PATCH = this.graph.PATCH<MessageDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<MessageDataModel>(this.path, this.query);
*/
}

export function _messages():Messages;
export function _messages(messageId:string):Message;
export function _messages(arg?:any):any {
    if (arg)
        return new Message(this.graph, this.path, arg);
    else
        return new Messages(this.graph, this.path);
}

export class Messages extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/messages/");
    }

    GetMessages = () => this.graph.GetCollection<MessageDataModel, Messages>(this.pathWithQuery, this, new Messages(this.graph));
/*
    CreateMessage = this.graph.POST<MessageDataModel>(this.path, this.query);
*/
}

export class Event extends Node {
    constructor(graph:Graph, path:string="", eventId:string) {
        super(graph, path + "/events/");
    }

    attachments = _attachments;

    GetEvent = () => this.graph.Get<EventDataModel, Event>(this.pathWithQuery, this);
/*
    PATCH = this.graph.PATCH<EventDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<EventDataModel>(this.path, this.query);
*/
}

export class Events extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/events/");
    }

    GetEvents = () => this.graph.GetCollection<EventDataModel, Events>(this.pathWithQuery, this, new Events(this.graph));
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

export function _events():Events;
export function _events(eventId:string):Event;
export function _events(arg?:any):any {
    if (arg)
        return new Event(this.graph, this.path, arg);
    else
        return new Events(this.graph, this.path);
}

export class CalendarView extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/calendarView");
    }

    GetCalendarView = () => this.graph.GetCollection<EventDataModel, CalendarView>(this.pathWithQuery, this, new CalendarView(this.graph));

    dateRange = (startDate:Date, endDate:Date) => this.odata(`startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`);
}

export class MailFolders extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/mailFolders");
    }

    GetMailFolders = () => this.graph.GetCollection<MailFolderDataModel, MailFolders>(this.pathWithQuery, this, new MailFolders(this.graph));
}

export class User extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, userId ? path + "/users/" + userId : path + "/me");
    }

    messages = _messages;
    events = _events;
    calendarView = () => new CalendarView(this.graph, this.path);
    mailFolders = () => new MailFolders(this.graph, this.path)

    GetUser = () => this.graph.Get<UserDataModel, User>(this.pathWithQuery, this); // REVIEW what about GetMe?
/*
    PATCH = this.graph.PATCH<UserDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<UserDataModel>(this.path, this.query);
*/
}

export class Users extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/users");
    }

    GetUsers = () => this.graph.GetCollection<UserDataModel, Users>(this.pathWithQuery, this, new Users(this.graph));
/*
    CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
*/
}
