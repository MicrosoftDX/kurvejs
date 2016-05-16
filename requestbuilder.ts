namespace Kurve {

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
        POST "/me/events/123/microsoft.graph.decline""

Graph operations are exposed through Promises:

    graph.me.messages
    .GetMessages()
    .then(messages =>
        messages.forEach(message =>
            console.log(message.subject)
        )
    )

All operations return a "_node" property which allows you to continue along the Graph path from the point where you left off:

    graph.me.messages.$("123").GetMessage().then(message =>
        console.log(message.subject);
        message._node.attachments.GetAttachments().then(attachments => // attachments._node === graph.me.messages.$("123").attachments
            attachments.forEach(attachment => 
                console.log(attachment.contentBytes)
            )
        )
    )

Members of returned collections also have a "_node" object:

        me.messages.GetMessages().then(messages =>
            messages.forEach(message =>
                message._node.attachments.GetAttachments().then(attachments =>
                    attachments.forEach(attachment =>
                        console.log(attachment.id);
                    )
                )
            )
        )

In this example message._node returns the expected node in the graph. But consider
the case of:

    me/directReports/{userId}
    
This returns a user object, but you can't do user operations. This operation is disallowed:

    me/directReports/{userId}/manager
    
Instead you must start again at the beginning:

    users/{userId}/manager
    
We facilitate this by setting the "_node" property on members of certain collections accordingly: 

        me.directReports.GetDirectReports().then(users =>
            users.forEach(user =>
                user._node.manager.GetManager()     // equivalent to users.$(user.id).manager.GetManager();
            )
        )

Operations which return paginated collections can return a "_next" request object. This can be utilized in a recursive function:

    ListMessageSubjects(messages:Collection<MessageDataModel, Messages, Message>) {
        messages.forEach(message => console.log(message.subject));
        if (messages._next)
            messages._next().then(nextMessages =>
                ListMessageSubjects(nextMessages)
            )
        })
    }
    graph.me.messages.GetMessages(new OData().select("subject")).then(messages =>
        ListMessageSubjects(messages)
    );
    
(With async/await support, an iteration pattern can be used intead of recursion)

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


export type Singleton<Model, N extends Node> = Model & {
    _node?: N,
    _item: Model
};

export function singletonFromResponse<Model, N extends Node>(response:any, node:N) {
    let singleton = response as Singleton<Model, N>;
    singleton._item = response as Model;
    singleton._node = node;
    return singleton;
}

export type ChildFactory<Model, N extends Node> = (id:string) => N;

export type Collection<Model, C extends CollectionNode, N extends Node> = Array<Singleton<Model, N>> & {
    _next?: () => Promise<Collection<Model, C, N>, Error>,
    _node?: C,
    _raw: any,
    _items: Model[]
};

export function collectionFromResponse<Model, C extends CollectionNode, N extends Node>(response:any, node:C, graph:Graph, childFactory?:ChildFactory<Model, N>, scopes?:string[]) {
    let collection = response.value as Collection<Model, C, N>;
    collection._node = node;
    collection._raw = response;
    collection._items = response.value as Model[];
    let nextLink = response["@odata.nextLink"];
    if (nextLink)
        collection._next = () => graph.GetCollection<Model, C, N>(nextLink, node, childFactory, scopes);
    if (childFactory)
        collection.forEach(item => item._node = item["id"] && childFactory(item["id"]));
    return collection;
}

export abstract class Node {
    constructor(protected graph:Graph, protected path:string) {
    }

    //Only adds scopes when linked to a v2 Oauth of kurve identity
    protected scopesForV2 = (scopes: string[]) =>
        this.graph.KurveIdentity && this.graph.KurveIdentity.getCurrentEndPointVersion() === EndPointVersion.v2 ? scopes : null;
    
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
    
    GetAttachments = (odataQuery?:ODataQuery) => this.graph.GetCollection<AttachmentDataModel, Attachments, Attachment>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2(Attachment.scopes[this.context]));
/*
    POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
*/
}

export class Message extends Node {
    constructor(graph:Graph, path:string="", messageId?:string) {
        super(graph, path + (messageId ? "/" + messageId : ""));
    }
    
    get attachments() { return new Attachments(this.graph, this.path, "messages"); }

    GetMessage  = (odataQuery?:ODataQuery) => this.graph.Get<MessageDataModel, Message>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.Read]));
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

    GetMessages     = (odataQuery?:ODataQuery) => this.graph.GetCollection<MessageDataModel, Messages, Message>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Mail.Read]));
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

    GetEvents = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, Events, Event>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Calendars.Read]));
/*
    POST = this.graph.POST<EventDataModel>(this.path, this.query);
*/
}

export class CalendarView extends CollectionNode {
    private static suffix = "/calendarView";
    constructor(graph:Graph, path:string="") {
        super(graph, path + CalendarView.suffix);
    }
    
    private $ = (eventId:string) => new Event(this.graph, this.path, eventId); // need to adjust this path
    
    dateRange = (startDate:Date, endDate:Date) => `startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`

    GetCalendarView = (odataQuery?:ODataQuery) => this.graph.GetCollection<EventDataModel, CalendarView, Event>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Calendars.Read]));
}


export class MailFolder extends Node {
    constructor(graph:Graph, path:string="", mailFolderId:string) {
        super(graph, path + (mailFolderId ? "/" + mailFolderId : ""));
    }

    GetMailFolder = (odataQuery?:ODataQuery) => this.graph.Get<MailFolderDataModel, MailFolder>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.Read]));
}

export class MailFolders extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/mailFolders");
    }

    $ = (mailFolderId:string) => new MailFolder(this.graph, this.path, mailFolderId);

    GetMailFolders = (odataQuery?:ODataQuery) => this.graph.GetCollection<MailFolderDataModel, MailFolders, MailFolder>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Mail.Read]));
}


export class Photo extends Node {    
    constructor(graph:Graph, path:string="", private context:string) {
        super(graph, path + "/photo" );
    }

    static scopes = {
        user: [Scopes.User.ReadBasicAll],
        group: [Scopes.Group.ReadAll],
        contact: [Scopes.Contacts.Read]
    }

    GetPhotoProperties = (odataQuery?:ODataQuery) => this.graph.Get<ProfilePhotoDataModel, Photo>(this.pathWithQuery(odataQuery), this, this.scopesForV2(Photo.scopes[this.context]));
    GetPhotoImage = (odataQuery?:ODataQuery) => this.graph.Get<any, any>(this.pathWithQuery(odataQuery, "/$value"), this, this.scopesForV2(Photo.scopes[this.context]), "blob");
}

export class Manager extends Node {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/manager" );
    }

    GetManager = (odataQuery?:ODataQuery) => this.graph.Get<UserDataModel, Manager>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.User.ReadAll]));
}

export class MemberOf extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/memberOf");
    }

    GetGroups = (odataQuery?:ODataQuery) => this.graph.GetCollection<GroupDataModel, MemberOf, Group>(this.pathWithQuery(odataQuery), this, Groups.$(this.graph), this.scopesForV2([Scopes.User.ReadAll]));
}

export class DirectReport extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, path + "/" + userId);
    }
    
    // seems like this should re-root its response under Users
    
    GetDirectReport = (odataQuery?:ODataQuery) => this.graph.Get<UserDataModel, DirectReport>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.User.Read]));
}
    
export class DirectReports extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/directReports");
    }

    $ = (userId:string) => new DirectReport(this.graph, this.path, userId);
    
    GetDirectReports = (odataQuery?:ODataQuery) => this.graph.GetCollection<UserDataModel, DirectReports, User>(this.pathWithQuery(odataQuery), this, Users.$(this.graph), this.scopesForV2([Scopes.User.Read]));
}

export class User extends Node {
    constructor(protected graph:Graph, path:string="", userId?:string) {
        super(graph, userId ? path + "/" + userId : path + "/me");
    }

    get messages()      { return new Messages(this.graph, this.path); }
    get events()        { return new Events(this.graph, this.path); }
    get calendarView()  { return new CalendarView(this.graph, this.path); }
    get mailFolders()   { return new MailFolders(this.graph, this.path) }
    get photo()         { return new Photo(this.graph, this.path, "user"); }
    get manager()       { return new Manager(this.graph, this.path); }
    get directReports() { return new DirectReports(this.graph, this.path); }
    get memberOf()      { return new MemberOf(this.graph, this.path); }

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
    
    static $ = (graph:Graph) => graph.users.$; 

    GetUsers = (odataQuery?:ODataQuery) => this.graph.GetCollection<UserDataModel, Users, User>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.User.Read]));
/*
    CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
*/
}

export class Group extends Node {
    constructor(protected graph:Graph, path:string="", groupId:string) {
        super(graph, path + "/" + groupId);
    }

    GetGroup = (odataQuery?:ODataQuery) => this.graph.Get<GroupDataModel, Group>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Group.ReadAll]));
}

export class Groups extends CollectionNode {
    constructor(graph:Graph, path:string="") {
        super(graph, path + "/groups");
    }
    
    $ = (groupId:string) => new Group(this.graph, this.path, groupId);
    
    static $ = (graph:Graph) => graph.groups.$;

    GetGroups = (odataQuery?:ODataQuery) => this.graph.GetCollection<GroupDataModel, Groups, Group>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Group.ReadAll]));
}


}