namespace Kurve {

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

    export type GraphObject<Model, N extends Node> = Model & {
        _context?: N,
        _item: Model
    };

    export type ChildFactory<Model, N extends Node> = (id:string) => N;

    export type GraphCollection<Model, C extends CollectionNode, N extends Node> = Array<GraphObject<Model, N>> & {
        _next?: () => Promise<GraphCollection<Model, C, N>, Error>,
        _context?: C,
        _raw: any,
        _items: Model[]
    };

    export abstract class Node {
        constructor(protected graph:Graph, protected path:string) {
        }

        //Only adds scopes when linked to a v2 Oauth of kurve identity
        protected scopesForV2 = (scopes: string[]) =>
            this.graph.KurveIdentity && this.graph.KurveIdentity.getCurrentEndPointVersion() === EndPointVersion.v2 ? scopes : null;
        
        pathWithQuery = (odataQuery?:ODataQuery, pathSuffix:string = "") => pathWithQuery(this.path + pathSuffix, odataQuery);
        
        protected graphObjectFromResponse = <Model, N extends Node>(response:any, node:N, childFactory?:ChildFactory<Model, N>) => {
            let singleton = response as GraphObject<Model, N>;
            singleton._item = response as Model;
            singleton._context = childFactory ? childFactory(singleton._item["id"]) : node;
            return singleton;
        }

        protected get<Model, N extends Node>(path:string, node:N, scopes?:string[], childFactory?:ChildFactory<Model, N>, responseType?:string): Promise<GraphObject<Model, N>, Error> {
            console.log("GET", path, scopes);
            var d = new Deferred<GraphObject<Model, N>, Error>();

            this.graph.get(path, (error, result) => {
                if (!responseType) {
                    var jsonResult = JSON.parse(result) ;

                    if (jsonResult.error) {
                        var errorODATA = new Error();
                        errorODATA.other = jsonResult.error;
                        d.reject(errorODATA);
                        return;
                    }
                    d.resolve(this.graphObjectFromResponse<Model, N>(jsonResult, node));
                } else {
                    d.resolve(this.graphObjectFromResponse<Model, N>(result, node));
                }

                
            }, responseType, scopes);

            return d.promise;
            }

        protected post<Model, N extends Node>(object:Model, path:string, node:N, scopes?:string[]): Promise<GraphObject<Model, N>, Error> {
            console.log("POST", path, scopes);
            var d = new Deferred<GraphObject<Model, N>, Error>();
            
    /*
            this.graph.post(object, path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Response<Model, N>({}, node));
            });
    */
            return d.promise;
            }

    }

    export abstract class CollectionNode extends Node {    
        private _nextLink:string;   // this is only set when the collection in question is from a nextLink

        pathWithQuery = (odataQuery?:ODataQuery, pathSuffix:string = "") => this._nextLink || pathWithQuery(this.path + pathSuffix, odataQuery);
        
        set nextLink(pathWithQuery:string) {
            this._nextLink = pathWithQuery;
        }
        
        protected graphCollectionFromResponse = <Model, C extends CollectionNode, N extends Node>(response:any, node:C, childFactory?:ChildFactory<Model, N>, scopes?:string[]) => {
            let collection = response.value as GraphCollection<Model, C, N>;
            collection._context = node;
            collection._raw = response;
            collection._items = response.value as Model[];
            let nextLink = response["@odata.nextLink"];
            if (nextLink)
                collection._next = () => this.getCollection<Model, C, N>(nextLink, node, childFactory, scopes);
            if (childFactory)
                collection.forEach(item => item._context = item["id"] && childFactory(item["id"]));
            return collection;
        }

        protected getCollection<Model, C extends CollectionNode, N extends Node>(path:string, node:C, childFactory:ChildFactory<Model, N>, scopes?:string[]): Promise<GraphCollection<Model, C, N>, Error> {
            console.log("GET collection", path, scopes);
            var d = new Deferred<GraphCollection<Model, C, N>, Error>();

            this.graph.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }
                
                d.resolve(this.graphCollectionFromResponse<Model, C, N>(jsonResult, node, childFactory, scopes));
            }, null, scopes);

            return d.promise;
            }
    }

    export class Attachment extends Node {
        constructor(graph:Graph, path:string="", private context:string, attachmentId?:string) {
            super(graph, path + (attachmentId ? "/" + attachmentId : ""));
        }

        static scopes = {
            messages: [Scopes.Mail.Read],
            events: [Scopes.Calendars.Read]
        }
        
        GetAttachment = (odataQuery?:ODataQuery) => this.get<AttachmentDataModel, Attachment>(this.pathWithQuery(odataQuery), this, this.scopesForV2(Attachment.scopes[this.context]));
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
        
        GetAttachments = (odataQuery?:ODataQuery) => this.getCollection<AttachmentDataModel, Attachments, Attachment>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2(Attachment.scopes[this.context]));
    /*
        POST = this.graph.POST<AttachmentDataModel>(this.path, this.query);
    */
    }

    export class Message extends Node {
        constructor(graph:Graph, path:string="", messageId?:string) {
            super(graph, path + (messageId ? "/" + messageId : ""));
        }
        
        get attachments() { return new Attachments(this.graph, this.path, "messages"); }

        GetMessage  = (odataQuery?:ODataQuery) => this.get<MessageDataModel, Message>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.Read]));
        SendMessage = (odataQuery?:ODataQuery) => this.post<MessageDataModel, Message>(null, this.pathWithQuery(odataQuery, "/microsoft.graph.sendMail"), this, this.scopesForV2([Scopes.Mail.Send]));
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

        GetMessages     = (odataQuery?:ODataQuery) => this.getCollection<MessageDataModel, Messages, Message>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Mail.Read]));
        CreateMessage   = (object:MessageDataModel, odataQuery?:ODataQuery) => this.post<MessageDataModel, Messages>(object, this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.ReadWrite]));
    }

    export class Event extends Node {
        constructor(graph:Graph, path:string="", eventId:string) {
            super(graph, path + (eventId ? "/" + eventId : ""));
        }

        get attachments() { return new Attachments(this.graph, this.path, "events"); }

        GetEvent = (odataQuery?:ODataQuery) => this.get<EventDataModel, Event>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Calendars.Read]));
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

        GetEvents = (odataQuery?:ODataQuery) => this.getCollection<EventDataModel, Events, Event>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Calendars.Read]));
    /*
        POST = this.graph.POST<EventDataModel>(this.path, this.query);
    */
    }

    export class CalendarView extends CollectionNode {
        constructor(graph:Graph, path:string="") {
            super(graph, path + "/calendarView");
        }
        
        private $ = (eventId:string) => new Event(this.graph, this.path, eventId); // need to adjust this path
        
        dateRange = (startDate:Date, endDate:Date) => `startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`

        GetEvents = (odataQuery?:ODataQuery) => this.getCollection<EventDataModel, CalendarView, Event>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Calendars.Read]));
    }


    export class MailFolder extends Node {
        constructor(graph:Graph, path:string="", mailFolderId:string) {
            super(graph, path + (mailFolderId ? "/" + mailFolderId : ""));
        }

        GetMailFolder = (odataQuery?:ODataQuery) => this.get<MailFolderDataModel, MailFolder>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Mail.Read]));
    }

    export class MailFolders extends CollectionNode {
        constructor(graph:Graph, path:string="") {
            super(graph, path + "/mailFolders");
        }

        $ = (mailFolderId:string) => new MailFolder(this.graph, this.path, mailFolderId);

        GetMailFolders = (odataQuery?:ODataQuery) => this.getCollection<MailFolderDataModel, MailFolders, MailFolder>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Mail.Read]));
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

        GetPhotoProperties = (odataQuery?:ODataQuery) => this.get<ProfilePhotoDataModel, Photo>(this.pathWithQuery(odataQuery), this, this.scopesForV2(Photo.scopes[this.context]));
        GetPhotoImage = (odataQuery?:ODataQuery) => this.get<any, any>(this.pathWithQuery(odataQuery, "/$value"), this, this.scopesForV2(Photo.scopes[this.context]), null, "blob");
    }

    export class Manager extends Node {
        constructor(graph:Graph, path:string="") {
            super(graph, path + "/manager" );
        }

        GetUser = (odataQuery?:ODataQuery) => this.get<UserDataModel, User>(this.pathWithQuery(odataQuery), null, this.scopesForV2([Scopes.User.ReadAll]), Users.$(this.graph));
    }

    export class MemberOf extends CollectionNode {
        constructor(graph:Graph, path:string="") {
            super(graph, path + "/memberOf");
        }

        GetGroups = (odataQuery?:ODataQuery) => this.getCollection<GroupDataModel, MemberOf, Group>(this.pathWithQuery(odataQuery), this, Groups.$(this.graph), this.scopesForV2([Scopes.User.ReadAll]));
    }

    export class DirectReport extends Node {
        constructor(protected graph:Graph, path:string="", userId?:string) {
            super(graph, path + "/" + userId);
        }
        
        GetUser = (odataQuery?:ODataQuery) => this.get<UserDataModel, User>(this.pathWithQuery(odataQuery), null, this.scopesForV2([Scopes.User.Read]), Users.$(this.graph));
    }
        
    export class DirectReports extends CollectionNode {
        constructor(graph:Graph, path:string="") {
            super(graph, path + "/directReports");
        }

        $ = (userId:string) => new DirectReport(this.graph, this.path, userId);
        
        GetUsers = (odataQuery?:ODataQuery) => this.getCollection<UserDataModel, DirectReports, User>(this.pathWithQuery(odataQuery), this, Users.$(this.graph), this.scopesForV2([Scopes.User.Read]));
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

        GetUser = (odataQuery?:ODataQuery) => this.get<UserDataModel, User>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.User.Read]));
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

        GetUsers = (odataQuery?:ODataQuery) => this.getCollection<UserDataModel, Users, User>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.User.Read]));
    /*
        CreateUser = this.graph.POST<UserDataModel>(this.path, this.query);
    */
    }

    export class Group extends Node {
        constructor(protected graph:Graph, path:string="", groupId:string) {
            super(graph, path + "/" + groupId);
        }

        GetGroup = (odataQuery?:ODataQuery) => this.get<GroupDataModel, Group>(this.pathWithQuery(odataQuery), this, this.scopesForV2([Scopes.Group.ReadAll]));
    }

    export class Groups extends CollectionNode {
        constructor(graph:Graph, path:string="") {
            super(graph, path + "/groups");
        }
        
        $ = (groupId:string) => new Group(this.graph, this.path, groupId);
        
        static $ = (graph:Graph) => graph.groups.$;

        GetGroups = (odataQuery?:ODataQuery) => this.getCollection<GroupDataModel, Groups, Group>(this.pathWithQuery(odataQuery), this, this.$, this.scopesForV2([Scopes.Group.ReadAll]));
    }

}