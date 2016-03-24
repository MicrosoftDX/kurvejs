
/*

GraphBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    graph.                      me, path, users
    graph.users(                two different options: with parameter returns a single user, without returns all users
    graph.me().                 events, messages, path
    
The API mirrors the structure of the Graph paths:
    to access:
        /users/billba@microsoft.com/messages/123-456-789/attachments
    use:
        graph.users("billba@microsoft.com").messages("123-456-789").attachments()

Different nodes surface appropriate functionality:
    graph.users("bill").        delete, events, get, messages, path
    graph.users().              get, path

Each node surfaces a path which can be passed to the channel of your choice to access the MS Graph: 
    graph.users("billba@microsoft.com").messages("123-456-789").attachments().pathWithQuery("$select=id,inline")
    => '/users/billba@microsoft.com/messages/123-456-789/attachments?$select=id,inline'

Or use the convenient built-in strongly-typed REST methods!
    graph.users("bill").messages().get()        => Graph.MessageDataModel[]
    graph.users("bill").event("123").get()      => Graph.EventDataModel

Each REST method has available to it the appropriate path via "this.pathWithQuery".
A similar approach would be used to convey scopes down the call chain.

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/

module GraphBuilder {
    
    export class Node {
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

    function attachments(): Attachments;
    function attachments(attachmentId:string): Attachment;
    function attachments(arg?:any):any {
        if (arg)
            return new Attachment(this.path + "/attachments/" + arg);
        else
            return new Attachments(this.path + "/attachments");
    }

    export class MessageDataModel {
        public id: string;
        public bodyPreview: string = "Sample body text";
    }

    export class Message extends Node {
        public get: GetObject<MessageDataModel> = getObject;
        public delete: DeleteObject<MessageDataModel> = deleteObject;
        public attachments = attachments;
        }

    export class Messages extends Node {
        public get: GetObject<MessageDataModel[]> = getObject;
    }

    function messages(): Messages;
    function messages(messageId:string): Message;
    function messages(arg?:any):any {
        if (arg)
            return new Message(this.path + "/messages/" + arg);
        else
            return new Messages(this.path + "/messages/");
    }

    export class EventDataModel {
        public id: string;
        public bodyPreview: string = "Sample body text";
    }

    export class Event extends Node {
        public get: GetObject<EventDataModel> = getObject;
        public delete: DeleteObject<EventDataModel> = deleteObject;
        public attachments = attachments;
        }

    export class Events extends Node {
        public get: GetObject<EventDataModel[]> = getObject;
    }

    function events(): Events;
    function events(eventId:string): Event;
    function events(arg?:any):any {
        if (arg)
            return new Event(this.path + "/events/" + arg);
        else
            return new Events(this.path + "/events/");
    }

    function calendarView(): Events;
    function calendarView(eventId:string): Event;
    function calendarView(arg?:any):any {
        if (arg)
            return new Event(this.path + "/calendarView/" + arg);
        else
            return new Events(this.path + "/calendarView/");
    }

    export class Me extends Node {
        public messages = messages;
        public events = events;
    }

    function me() {
        return new Me("/me");
    }

    export class UserDataModel {
        public id: string;
        public name: string = "Bill Barnes";
    }

    export class User extends Node {
        public get: GetObject<UserDataModel> = getObject;
        public delete: DeleteObject<UserDataModel> = deleteObject;
        public messages = messages;
        public events = events;
    }

    export class Users extends Node {
        public get: GetObject<UserDataModel[]> = getObject;
    }

    function users(): Users;
    function users(userId:string): User;
    function users(arg?:any):any {
        if (arg)
            return new User("/users/" + arg);
        else
            return new Users("/users/");
    }

    export class Root {
        public me = me;
        public users = users;
    }
}

var graph = new GraphBuilder.Root();

// Explore the graph by typing below here!

