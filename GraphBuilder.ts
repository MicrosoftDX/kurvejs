
/*

GraphBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    graph.                      me, path, users
    graph.me                    event, events, message, messages, path
    graph.me.event              event(eventId:string) => Event
    
The API mirrors the structure of the Graph paths:
    to access:
        /users/billba@microsoft.com/messages/123-456-789/attachments
    instead of the Kurve-y way of doing things:
        graph.messageAttachmentsForUser("billba@microsoft.com", "123-456-789")
    we do:
        graph.user("billba@microsoft.com").message("123-456-789").attachments

Different nodes surface appropriate functionality:
    graph.user("bill").         delete, event, events, get, message, messages, path
    graph.users.                get, path

Each node surfaces a path which can be passed to the channel of your choice to access the MS Graph: 
    graph.user("billba@microsoft.com").message("123-456-789").attachments.pathWithQuery("$select=id,inline")
    => '/users/billba@microsoft.com/messages/123-456-789/attachments?$select=id,inline'

Or use the convenient built-in strongly-typed REST methods!
    graph.user("bill").messages().get()        => Graph.MessageDataModel[]
    graph.user("bill").event("123").get()      => Graph.EventDataModel

Each REST method has available to it the appropriate path via "this.pathWithQuery".
A similar approach would be used to convey scopes down the call chain.

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/

module GraphBuilder {
    
    abstract class Node {
        constructor(public path: string) {}
        public pathWithQuery = (odataQuery?:string) => this.path + (odataQuery ? "?" + odataQuery : "");
    }  

    interface GetObject<T> {
        (odataQuery?:string): T;
    }

    function getObject<T>(odataQuery?:string): T {
        console.log("path", this.pathWithQuery(odataQuery));
        return {} as T;
    }

    interface DeleteObject<T> {
        (odataQuery?:string);
    }

    function deleteObject<T>(odataQuery?:string) {
        console.log("path", this.pathWithQuery(odataQuery));
        return;
    }

    interface PostObject<T> {
        (odataQuery?:string);
    }

    function postObject<T>(odataQuery?:string) {
        console.log("path", this.pathWithQuery(odataQuery));
        return;
    }

    export class AttachmentDataModel {
        public id: string;
        public isInline: boolean;
    }

    export class Attachment extends Node {
        public get: GetObject<AttachmentDataModel> = getObject;
        public delete: DeleteObject<AttachmentDataModel> = deleteObject;
    }

    export class Attachments extends Node {
        public get: GetObject<AttachmentDataModel[]> = getObject;
    }

    abstract class HasAttachments extends Node {
        public attachment = (attachmentId:string) => new Attachment(this.path + "/attachments/" + attachmentId);
        public attachments = new Attachments(this.path + "/attachments");   
    }

    export class MessageDataModel {
        public id: string;
        public bodyPreview: string = "Sample body text";
    }

    export class Message extends HasAttachments {
        public get: GetObject<MessageDataModel> = getObject;
        public delete: DeleteObject<MessageDataModel> = deleteObject;
        }

    export class Messages extends Node {
        public get: GetObject<MessageDataModel[]> = getObject;
    }
    
    export class EventDataModel {
        public id: string;
        public bodyPreview: string = "Sample body text";
    }

    export class Event extends HasAttachments {
        public get: GetObject<EventDataModel> = getObject;
        public delete: DeleteObject<EventDataModel> = deleteObject;
        }

    export class Events extends Node {
        public get: GetObject<EventDataModel[]> = getObject;
    }

    export class Me extends Node {
        public message = (messageId: string) => new Message(this.path + "/messages/" + messageId);
        public messages = new Messages(this.path + "/messages");
        public event = (eventId: string) => new Event(this.path + "/events/" + eventId);
        public events = new Events(this.path + "/events");
        public calendarView = new Events(this.path + "/calendarView");
    }

    export class UserDataModel {
        public id: string;
        public name: string = "Bill Barnes";
    }

    export class User extends Me {
        public get: GetObject<UserDataModel> = getObject;
        public delete: DeleteObject<UserDataModel> = deleteObject;
    }

    export class Users extends Node {
        public get: GetObject<UserDataModel[]> = getObject;
    }

    export class Root {
        public me = new Me("/me");
        public user = (userId:string) => new User("/users/" + userId);
        public users = new Users("/users/");
    }
}

var graph = new GraphBuilder.Root();

// Explore the graph by typing below here!
