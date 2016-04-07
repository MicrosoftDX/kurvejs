/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing and see how intellisense helps you explore the graph:
    graph.                          me, users
    graph.me().                     events, messages, calendarView, mailFolders, GetUser, odata, select, ...
    graph.me().events().            GetEvents, id, odata, select, ...
    graph.me().events().id          (eventId:string) => Event
    graph.me().events().id("123").  GetEvent, odata, select, ...

Each endpoint exposes the set of available Graph operations through strongly typed methods:
    graph.me().GetUser()
        GET "/me"" => UserDataModel
    graph.me().events().GetEvents()
        GET "/me/events"" => EventDataModel[]
    graph.me().events().CreateEvent(event:EventDataModel) 
        POST EventDataModel => "/me/events" 
        
Graph operations are exposed through Promises:
    graph.me().messages()
    .GetMessages()
    .then(collection =>
        collection.objects.forEach(message =>
            console.log(message.subject)
        )
    )

All operations return a "self" property which allows you to continue along the Graph path from the point where you left off:
    graph.me().messages().id("123").GetMessage().then(response =>
        console.log(response.object.subject);
        response.self.attachments().GetAttachments().then(collection => // response.self === graph.me().messages().id("123")
            collection.objects.forEach(attachment => 
                console.log(attachment.contentBytes)
            )
        )
    )

Operations which return paginated collections can return a "next" request object. This can be utilized in a recursive function:
    ListMessageSubjects(messages:Messages) {
        messages.GetMessages().then(collection => {
            collection.objects.forEach(message => console.log(message.subject));
            if (collection.next)
                ListMessageSubjects(collection.next);
        })
    }
    ListMessageSubjects(graph.me().messages());

Endpoints can be decorated with ODATA helpers, which can be chained:
    graph.me().messages().select("subject", "id")
        /me/messages/$select=subject,id
    graph.me().messages().select("subject", "id").orderby("id")
        /me/messages/$select=subject,id&$orderby=id

Or live close to the metal by writing your own ODATA directly:
    graph.me().messages().odata("$select=subject,id&$orderby=id")

These helpers are a little quirky.

Quirk 1: ODATA helpers are only applied if they are at the end of the request:
    graph.me().messages().select("subject", "id")
        /me/messages/$select=subject,id
    graph.me().select("name").messages()
        /me/messages/
    graph.me().select("name").messages().select("subject", "id")
        /me/messages/$select=subject,id

This shows up (in a handy way) when you reuse requests stored in variables:
    let message = graph.me().messages().id("123").select("subject", "id")
    message.GetMessage().then(...)
    message.attachments().select("contentBytes").GetAttachments().then(...) // message's ODATA is ignored

Quirk 2: ODATA helpers change the object they decorate:
    let foo = graph.me()
        foo.pathWithQuery => "/me"
    foo.select("name").GetUser()
        foo.pathWithQuery => "/me?$select=name"
        
[I can't figure out how to fix this. An ugly workaround is to add a helper to revert the object, e.g. foo.clearQuery()]

This initial stab only includes a few familiar pieces of the Microsoft Graph.
*/

import { Promise } from "./promises";
import { Graph } from "./graph";
import { Error } from "./identity";
import { UserDataModel, AttachmentDataModel, MessageDataModel, EventDataModel, MailFolderDataModel } from './models';

export class Singleton<Model, N extends Node> {
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
        super(graph, path + "/" + attachmentId);
    }

    GetAttachment = () => this.graph.Get<AttachmentDataModel, Attachment>(this.pathWithQuery, this);
/*    
    PATCH = this.graph.PATCH<AttachmentDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<AttachmentDataModel>(this.path, this.query);
*/
}

export class Attachments extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/attachments");
    }

    id = (attachmentId:string) => new Attachment(this.graph, this.path, attachmentId);

    GetAttachments = () => this.graph.GetCollection<AttachmentDataModel, Attachments>(this.pathWithQuery, this, new Attachments(this.graph));
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

export class Message extends Node {
    constructor(graph:Graph, path:string="", messageId:string) {
        super(graph, path + "/" + messageId);
    }
    
    attachments = () => new Attachments(this.graph, this.path);

    GetMessage = () => this.graph.Get<MessageDataModel, Message>(this.pathWithQuery, this);
    SendMessage = () => this.graph.Post<MessageDataModel, Message>(null, pathWithQuery(this.path + "/microsoft.graph.sendMail", this.query), this);
/*
    PATCH = this.graph.PATCH<MessageDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<MessageDataModel>(this.path, this.query);
*/
}

export class Messages extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/messages");
    }

    id = (messageId:string) => new Message(this.graph, this.path, messageId);

    GetMessages = () => this.graph.GetCollection<MessageDataModel, Messages>(this.pathWithQuery, this, new Messages(this.graph));
    CreateMessage = (object:MessageDataModel) => this.graph.Post<MessageDataModel, Messages>(object, this.pathWithQuery, this);
}

export class Event extends Node {
    constructor(graph:Graph, path:string="", eventId:string) {
        super(graph, path + "/" + eventId);
    }

    attachments = () => new Attachments(this.graph, this.path);

    GetEvent = () => this.graph.Get<EventDataModel, Event>(this.pathWithQuery, this);
/*
    PATCH = this.graph.PATCH<EventDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<EventDataModel>(this.path, this.query);
*/
}

export class Events extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/events");
    }

    id = (eventId:string) => new Event(this.graph, this.path, eventId);

    GetEvents = () => this.graph.GetCollection<EventDataModel, Events>(this.pathWithQuery, this, new Events(this.graph));
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

export class CalendarView extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/calendarView");
    }

    dateRange = (startDate:Date, endDate:Date) => this.odata(`startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`);

    GetCalendarView = () => this.graph.GetCollection<EventDataModel, CalendarView>(this.pathWithQuery, this, new CalendarView(this.graph));
}

export class MailFolder extends Node {
    constructor(graph:Graph, path:string="", mailFolderId:string) {
        super(graph, path + "/" + mailFolderId);
    }

    GetMailFolder = () => this.graph.Get<MailFolderDataModel, MailFolder>(this.pathWithQuery, this);
}

export class MailFolders extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/mailFolders");
    }

    id = (mailFolderId:string) => new MailFolder(this.graph, this.path, mailFolderId);

    GetMailFolders = () => this.graph.GetCollection<MailFolderDataModel, MailFolders>(this.pathWithQuery, this, new MailFolders(this.graph));
}

export class User extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, userId ? path + "/" + userId : path + "/me");
    }

    messages = () => new Messages(this.graph, this.path);
    events = () => new Events(this.graph, this.path);
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

    id = (userId:string) => new User(this.graph, this.path, userId);

    GetUsers = () => this.graph.GetCollection<UserDataModel, Users>(this.pathWithQuery, this, new Users(this.graph));
/*
    CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
*/
}
