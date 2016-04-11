/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing and see how intellisense helps you explore the graph:
    graph.                      me, users
    graph.me.                   events, messages, calendarView, mailFolders, GetUser, odata, select, ...
    graph.me.events.            GetEvents, $, odata, select, ...
    graph.me.events.$           (eventId:string) => Event
    graph.me.events.$("123").   GetEvent, odata, select, ...

Each endpoint exposes the set of available Graph operations through strongly typed methods:
    graph.me.GetUser() => UserDataModel
        GET "/me"
    graph.me.events.GetEvents => EventDataModel[]
        GET "/me/events"
    graph.me.events.CreateEvent(event:EventDataModel) 
        POST "/me/events"

Graph operations are exposed through Promises:
    graph.me.messages
    .GetMessages()
    .then(collection =>
        collection.objects.forEach(message =>
            console.log(message.subject)
        )
    )

All operations return a "self" property which allows you to continue along the Graph path from the point where you left off:
    graph.me.messages.$("123").GetMessage().then(response =>
        console.log(response.object.subject);
        response.self.attachments.GetAttachments().then(collection => // response.self === graph.me.messages.$("123")
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
    ListMessageSubjects(graph.me.messages);

Every Graph operation may include OData queries:
    graph.me.messages.GetMessages("$select=subject,id&$orderby=id")
        /me/messages/$select=subject,id&$orderby=id

There is an optional OData helper to aid in constructing more complex queries:
    graph.me.messages.GetMessages(new OData()
        .select("subject", "id")
        .orderby("id")
    )
        /me/messages/$select=subject,id&$orderby=id

Some operations include parameters which transform into OData queries 
    graph.me.calendarView.GetCalendarView([start],[end], [odataQuery])
        /me/calendarView?startDateTime=[start]&endDateTime=[end]&[odataQuery]

Note: This initial stab only includes a few familiar pieces of the Microsoft Graph.
*/

import { Promise } from "./promises";
import { Graph } from "./graph";
import { Error } from "./identity";
import { UserDataModel, AttachmentDataModel, MessageDataModel, EventDataModel, MailFolderDataModel } from './models';

let queryUnion = (query1:string, query2:string) => (query1 ? query1 + (query2 ? "&" + query2 : "" ) : query2); 

export class OData {
    constructor(public query?:string) {
    }
    
    toString = () => this.query;

    odata = (query:string) => {
        this.query = queryUnion(this.query, query);
        return this;
    }

    select   = (...fields:string[])  => this.odata(`$select=${fields.join(",")}`);
    expand   = (...fields:string[])  => this.odata(`$expand=${fields.join(",")}`);
    filter   = (query:string)        => this.odata(`$filter=${query}`);
    orderby  = (...fields:string[])  => this.odata(`$orderby=${fields.join(",")}`);
    top      = (items:Number)        => this.odata(`$top=${items}`);
    skip     = (items:Number)        => this.odata(`$skip=${items}`);
}

type ODataQuery = OData | string;

let pathWithQuery = (path:string, odataQuery?:ODataQuery) => {
    let query = odataQuery && odataQuery.toString();
    return path + (query ? "?" + query : "");
}

export class Singleton<Model, N extends Node> {
    constructor(public raw:any, public self:N) {
    }

    get object() {
        return this.raw as Model;
    }
}

export class Collection<Model, N extends CollectionNode> {
    constructor(public raw:any, public self:N, public next:N) {
        let nextLink = this.raw["@odata.nextLink"];
        if (nextLink) {
            this.next.nextLink = nextLink;
        } else {
            this.next = undefined;
        }
    }

    get objects() {
        return (this.raw.value ? this.raw.value : [this.raw]) as Model[];
    }
}

export abstract class Node {
    constructor(protected graph:Graph, protected path:string) {
    }

    pathWithQuery = (odataQuery?:ODataQuery, pathSuffix:string = "") => pathWithQuery(this.path + pathSuffix, odataQuery);
}

export abstract class CollectionNode extends Node {    
    private _nextLink:string;   // this is only set when the collection in question is from a nextLink

    pathWithQuery = (odataQuery?:ODataQuery, pathSuffix:string = "") => this._nextLink || pathWithQuery(this.path + pathSuffix, odataQuery);
    
    set nextLink(pathWithQuery:string) {
        this._nextLink = pathWithQuery;
    }
}

export class Attachment extends Node {
    constructor(graph:Graph, path:string="", attachmentId?:string) {
        super(graph, path + (attachmentId ? "/" + attachmentId : ""));
    }

    GetAttachment = (odataQuery?:ODataQuery) => this.graph.Get<AttachmentDataModel, Attachment>(this.pathWithQuery(odataQuery), this, null);
/*    
    PATCH = this.graph.PATCH<AttachmentDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<AttachmentDataModel>(this.path, this.query);
*/
}

export class Attachments extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/attachments");
    }

    $ = (attachmentId:string) => new Attachment(this.graph, this.path, attachmentId);

    GetAttachments = (odataQuery?:ODataQuery) => this.graph.GetCollection<AttachmentDataModel, Attachments>(this.pathWithQuery(odataQuery), this, new Attachments(this.graph));
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

export class Message extends Node {
    constructor(graph:Graph, path:string="", messageId?:string) {
        super(graph, path + (messageId ? "/" + messageId : ""));
    }
    
    get attachments() { return new Attachments(this.graph, this.path); }

    GetMessage  = (odataQuery?:ODataQuery) => this.graph.Get<MessageDataModel, Message>(this.pathWithQuery(odataQuery), this);
    SendMessage = (odataQuery?:ODataQuery) => this.graph.Post<MessageDataModel, Message>(null, this.pathWithQuery(odataQuery, "/microsoft.graph.sendMail"), this);
/*
    PATCH = this.graph.PATCH<MessageDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<MessageDataModel>(this.path, this.query);
*/
}

export class Messages extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/messages");
    }

    $ = (messageId:string) => new Message(this.graph, this.path, messageId);

    GetMessages     = (odataQuery?:ODataQuery) => this.graph.GetCollection<MessageDataModel, Messages>(this.pathWithQuery(odataQuery), this, new Messages(this.graph));
    CreateMessage   = (object:MessageDataModel, odataQuery?:ODataQuery) => this.graph.Post<MessageDataModel, Messages>(object, this.pathWithQuery(odataQuery), this);
}

export class Event extends Node {
    constructor(graph:Graph, path:string="", eventId:string) {
        super(graph, path + (eventId ? "/" + eventId : ""));
    }

    get attachments() { return new Attachments(this.graph, this.path); }

    GetEvent = (odataQuery?:ODataQuery) => this.graph.Get<EventDataModel, Event>(this.pathWithQuery(odataQuery), this);
/*
    PATCH = this.graph.PATCH<EventDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<EventDataModel>(this.path, this.query);
*/
}

export class Events extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/events");
    }

    $ = (eventId:string) => new Event(this.graph, this.path, eventId);

    GetEvents = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, Events>(this.pathWithQuery(odataQuery), this, new Events(this.graph));
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

export class CalendarView extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/calendarView");
    }

    GetCalendarView = (startDate?:Date, endDate?:Date, odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, CalendarView>(this.pathWithQuery(queryUnion(`startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`, odataQuery && odataQuery.toString())), this, new CalendarView(this.graph));
}

export class MailFolder extends Node {
    constructor(graph:Graph, path:string="", mailFolderId:string) {
        super(graph, path + (mailFolderId ? "/" + mailFolderId : ""));
    }

    GetMailFolder = (odataQuery?:ODataQuery) => this.graph.Get<MailFolderDataModel, MailFolder>(this.pathWithQuery(odataQuery), this);
}

export class MailFolders extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/mailFolders");
    }

    $ = (mailFolderId:string) => new MailFolder(this.graph, this.path, mailFolderId);

    GetMailFolders = (odataQuery?:ODataQuery) => this.graph.GetCollection<MailFolderDataModel, MailFolders>(this.pathWithQuery(odataQuery), this, new MailFolders(this.graph));
}

export class User extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, userId ? path + "/" + userId : path + "/me");
    }

    get messages()      { return new Messages(this.graph, this.path); }
    get events()        { return new Events(this.graph, this.path); }
    get calendarView()  { return new CalendarView(this.graph, this.path); }
    get mailFolders()   { return new MailFolders(this.graph, this.path) }

    GetUser = (odataQuery?:ODataQuery) => this.graph.Get<UserDataModel, User>(this.pathWithQuery(odataQuery), this); // Sorry, no GetMe()
/*
    PATCH = this.graph.PATCH<UserDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<UserDataModel>(this.path, this.query);
*/
}

export class Users extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/users");
    }

    $ = (userId:string) => new User(this.graph, this.path, userId);

    GetUsers = (odataQuery:ODataQuery) => this.graph.GetCollection<UserDataModel, Users>(this.pathWithQuery(odataQuery), this, new Users(this.graph));
/*
    CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
*/
}

