module kurve {

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

Certain Graph endpoints are implemented as OData "Functions". These are not treated as Graph nodes. They're just methods: 

    graph.me.events.$("123").DeclineEvent(eventResponse:EventResponse)
        POST "/me/events/123/microsoft.graph.decline

Graph operations are exposed through Promises:

    graph.me.messages
    .GetMessages()
    .then(collection =>
        collection.items.forEach(message =>
            console.log(message.subject)
        )
    )

All operations return a "self" property which allows you to continue along the Graph path from the point where you left off:

    graph.me.messages.$("123").GetMessage().then(singleton =>
        console.log(singleton.item.subject);
        singleton.self.attachments.GetAttachments().then(collection => // singleton.self === graph.me.messages.$("123")
            collection.items.forEach(attachment => 
                console.log(attachment.contentBytes)
            )
        )
    )

Operations which return paginated collections can return a "next" request object. This can be utilized in a recursive function:

    ListMessageSubjects(messages:Messages) {
        messages.GetMessages().then(collection => {
            collection.items.forEach(message => console.log(message.subject));
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


export class Scopes {
    private static rootUrl = "https://graph.microsoft.com/";
    static General = {
        OpenId: "openid",
        OfflineAccess: "offline_access",
    }
    static User = {
        Read: Scopes.rootUrl + "User.Read",
        ReadAll: Scopes.rootUrl + "User.Read.All",
        ReadWrite: Scopes.rootUrl + "User.ReadWrite",
        ReadWriteAll: Scopes.rootUrl + "User.ReadWrite.All",
        ReadBasicAll: Scopes.rootUrl + "User.ReadBasic.All",
    }
    static Contacts = {
        Read: Scopes.rootUrl + "Contacts.Read",
        ReadWrite: Scopes.rootUrl + "Contacts.ReadWrite",
    }
    static Directory = {
        ReadAll: Scopes.rootUrl + "Directory.Read.All",
        ReadWriteAll: Scopes.rootUrl + "Directory.ReadWrite.All",
        AccessAsUserAll: Scopes.rootUrl + "Directory.AccessAsUser.All",
    }
    static Group = {
        ReadAll: Scopes.rootUrl + "Group.Read.All",
        ReadWriteAll: Scopes.rootUrl + "Group.ReadWrite.All",
        AccessAsUserAll: Scopes.rootUrl + "Directory.AccessAsUser.All"
    }
    static Mail = {
        Read: Scopes.rootUrl + "Mail.Read",
        ReadWrite: Scopes.rootUrl + "Mail.ReadWrite",
        Send: Scopes.rootUrl + "Mail.Send",
    }
    static Calendars = {
        Read: Scopes.rootUrl + "Calendars.Read",
        ReadWrite: Scopes.rootUrl + "Calendars.ReadWrite",
    }
    static Files = {
        Read: Scopes.rootUrl + "Files.Read",
        ReadAll: Scopes.rootUrl + "Files.Read.All",
        ReadWrite: Scopes.rootUrl + "Files.ReadWrite",
        ReadWriteAppFolder: Scopes.rootUrl + "Files.ReadWrite.AppFolder",
        ReadWriteSelected: Scopes.rootUrl + "Files.ReadWrite.Selected",
    }
    static Tasks = {
        ReadWrite: Scopes.rootUrl + "Tasks.ReadWrite",
    }
    static People = {
        Read: Scopes.rootUrl + "People.Read",
        ReadWrite: Scopes.rootUrl + "People.ReadWrite",
    }
    static Notes = {
        Create: Scopes.rootUrl + "Notes.Create",
        ReadWriteCreatedByApp: Scopes.rootUrl + "Notes.ReadWrite.CreatedByApp",
        Read: Scopes.rootUrl + "Notes.Read",
        ReadAll: Scopes.rootUrl + "Notes.Read.All",
        ReadWriteAll: Scopes.rootUrl + "Notes.ReadWrite.All",
    }
}

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

    get item() {
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

    get items() {
        return (this.raw.value ? this.raw.value : [this.raw]) as Model[];
    }
}

export abstract class Node {
    constructor(protected graph:Graph, protected path:string) {
    }

    //Only adds scopes when linked to a v2 Oauth of kurve identity
    protected scopesForV2 = (scopes: string[]) =>
        this.graph.KurveIdentity && this.graph.KurveIdentity.getCurrentOauthVersion() === OAuthVersion.v2 ? scopes : null;
    
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
    constructor(graph:Graph, path:string="", private context:string, attachmentId?:string) {
        super(graph, path + (attachmentId ? "/" + attachmentId : ""));
    }

    static scopes = {
        messages: [Scopes.Mail.Read],
        events: [Scopes.Calendars.Read],
    }

    GetAttachment = (odataQuery?:ODataQuery) => this.graph.Get<AttachmentDataModel, Attachment>(this.pathWithQuery(odataQuery), this, this.scopesForV2(Attachment.scopes[this.context]));
/*    
    PATCH = this.graph.PATCH<AttachmentDataModel>(this.path, this.query);
    DELETE = this.graph.DELETE<AttachmentDataModel>(this.path, this.query);
*/
}

export class Attachments extends CollectionNode {
    constructor(graph:Graph, path:string="", private context:string) {
        super(graph, path + "/attachments");
    }

    $ = (attachmentId:string) => new Attachment(this.graph, this.path, this.context, attachmentId);
    
    GetAttachments = (odataQuery?:ODataQuery) => this.graph.GetCollection<AttachmentDataModel, Attachments>(this.pathWithQuery(odataQuery), this, new Attachments(this.graph, null, this.context), this.scopesForV2(Attachment.scopes[this.context]));
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

export class Message extends Node {
    constructor(graph:Graph, path:string="", messageId?:string) {
        super(graph, path + (messageId ? "/" + messageId : ""));
    }
    
    get attachments() { return new Attachments(this.graph, this.path, "messages"); }

    GetMessage  = (odataQuery?:ODataQuery) => this.graph.Get<MessageDataModel, Message>(this.pathWithQuery(odataQuery), this, [Scopes.Mail.Read]);
    SendMessage = (odataQuery?:ODataQuery) => this.graph.Post<MessageDataModel, Message>(null, this.pathWithQuery(odataQuery, "/microsoft.graph.sendMail"), this, this.scopesForV2([Scopes.Mail.Send]));
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

    GetMessages     = (odataQuery?:ODataQuery) => this.graph.GetCollection<MessageDataModel, Messages>(this.pathWithQuery(odataQuery), this, new Messages(this.graph), this.scopesForV2([Scopes.Mail.Read]));
    CreateMessage   = (object:MessageDataModel, odataQuery?:ODataQuery) => this.graph.Post<MessageDataModel, Messages>(object, this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.ReadWrite]));
}

export class Event extends Node {
    constructor(graph:Graph, path:string="", eventId:string) {
        super(graph, path + (eventId ? "/" + eventId : ""));
    }

    get attachments() { return new Attachments(this.graph, this.path, "events"); }

    GetEvent = (odataQuery?:ODataQuery) => this.graph.Get<EventDataModel, Event>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Calendars.Read]));
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

    GetEvents = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, Events>(this.pathWithQuery(odataQuery), this, new Events(this.graph), this.scopesForV2([Scopes.Calendars.Read]));
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

export class CalendarView extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/calendarView");
    }
    
    dateRange = (startDate:Date, endDate:Date) => `startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`

    GetCalendarView = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, CalendarView>(this.pathWithQuery(odataQuery), this, new CalendarView(this.graph), [Scopes.Calendars.Read]);
}


export class MailFolder extends Node {
    constructor(graph:Graph, path:string="", mailFolderId:string) {
        super(graph, path + (mailFolderId ? "/" + mailFolderId : ""));
    }


    GetMailFolder = (odataQuery?:ODataQuery) => this.graph.Get<MailFolderDataModel, MailFolder>(this.pathWithQuery(odataQuery), this, [Scopes.Mail.Read]);
}

export class MailFolders extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/mailFolders");
    }

    $ = (mailFolderId:string) => new MailFolder(this.graph, this.path, mailFolderId);

    GetMailFolders = (odataQuery?:ODataQuery) => this.graph.GetCollection<MailFolderDataModel, MailFolders>(this.pathWithQuery(odataQuery), this, new MailFolders(this.graph), this.scopesForV2([Scopes.Mail.Read]));
}

//let usersScopes = [Scopes.User.ReadBasicAll, Scopes.User.ReadAll, Scopes.User.ReadWriteAll, Scopes.Directory.ReadAll, Scopes.Directory.ReadWriteAll, Scopes.Directory.AccessAsUserAll];

export class User extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, userId ? path + "/" + userId : path + "/me");
    }

    get messages()      { return new Messages(this.graph, this.path); }
    get events()        { return new Events(this.graph, this.path); }
    get calendarView()  { return new CalendarView(this.graph, this.path); }
    get mailFolders()   { return new MailFolders(this.graph, this.path) }

    GetUser = (odataQuery?:ODataQuery) => this.graph.Get<UserDataModel, User>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.User.Read]));
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

    GetUsers = (odataQuery?:ODataQuery) => this.graph.GetCollection<UserDataModel, Users>(this.pathWithQuery(odataQuery), this, new Users(this.graph), [Scopes.User.Read]);
/*
    CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
*/
}



} //remove during bundling