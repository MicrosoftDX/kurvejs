/*

RequestBuilder allows you to discover and access the Microsoft Graph using Visual Studio Code intellisense.

Just start typing at the bottom of this file and see how intellisense helps you explore the graph:
    root.                      me, path, users, endpoints
    root.me                    event, events, message, messages, path
    root.me.event              event(eventId:string) => Event
    
The API mirrors the structure of the Graph paths:
    to access:
        /users/billba@microsoft.com/messages/123-456-789/attachments
    instead of the Kurve-y way of doing things:
        graph.messageAttachmentsForUser("billba@microsoft.com", "123-456-789")
    we do:
        root.user("billba@microsoft.com").message("123-456-789").attachments

Different nodes surface appropriate functionality:
    root.user("bill").         delete, event, events, get, message, messages, path
    root.users.                get, path

Each node surfaces a path which can be passed to the channel of your choice to access the MS Graph:
    root.user("billba@microsoft.com").message("123-456-789").attachments.pathWithQuery("$select=id,inline")
    => '/users/billba@microsoft.com/messages/123-456-789/attachments?$select=id,inline'

Or use the convenient built-in strongly-typed REST methods!
    root.user("bill").messages().get()        => Graph.MessageDataModel
    root.user("bill").event("123").get()      => Graph.EventDataModel

Each REST method has available to it the appropriate path via "this.pathWithQuery".
A similar approach would be used to convey scopes down the call chain.

In this proof-of-concept the path building works, but the REST methods are stubs, and greatly simplified ones at that.
In a real version we'd add Async versions and incorporate identity handling.

Finally this initial stab only includes a few familiar pieces of the Microsoft Graph.
However I have examined the 1.0 and Beta docs closely and I believe that this approach is extensible to the full graph.

*/
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
// A mock of Kurve for testing purposes.
var Kurve;
(function (Kurve) {
    var UserDataModel = (function () {
        function UserDataModel() {
            this.userField = "I am a user.";
        }
        return UserDataModel;
    }());
    Kurve.UserDataModel = UserDataModel;
    var MessageDataModel = (function () {
        function MessageDataModel() {
            this.messageField = "I am a message.";
        }
        return MessageDataModel;
    }());
    Kurve.MessageDataModel = MessageDataModel;
    var EventDataModel = (function () {
        function EventDataModel() {
            this.eventField = "I am an event.";
        }
        return EventDataModel;
    }());
    Kurve.EventDataModel = EventDataModel;
    var AttachmentDataModel = (function () {
        function AttachmentDataModel() {
            this.attachmentField = "I am an attachment.";
        }
        return AttachmentDataModel;
    }());
    Kurve.AttachmentDataModel = AttachmentDataModel;
})(Kurve || (Kurve = {}));
var RequestBuilder;
(function (RequestBuilder) {
    var Action = (function () {
        function Action(pathWithQuery, scopes) {
            this.pathWithQuery = pathWithQuery;
            this.scopes = scopes;
        }
        return Action;
    }());
    RequestBuilder.Action = Action;
    var Node = (function () {
        function Node(path, query) {
            if (path === void 0) { path = ""; }
            this.path = path;
            this.query = query;
        }
        Object.defineProperty(Node.prototype, "pathWithQuery", {
            get: function () { return this.path + (this.query ? "?" + this.query : ""); },
            enumerable: true,
            configurable: true
        });
        return Node;
    }());
    RequestBuilder.Node = Node;
    function queryUnion(query1, query2) {
        if (query1)
            return query1 + (query2 ? "&" + query2 : "");
        else
            return query2;
    }
    var AddQuery = (function (_super) {
        __extends(AddQuery, _super);
        function AddQuery(path, query, actions) {
            if (path === void 0) { path = ""; }
            _super.call(this, path, query);
            this.path = path;
            this.query = query;
            if (actions) {
                this.actions = actions;
                for (var verb in actions)
                    actions[verb].pathWithQuery = this.pathWithQuery;
            }
        }
        return AddQuery;
    }(Node));
    RequestBuilder.AddQuery = AddQuery;
    var NodeWithQuery = (function (_super) {
        __extends(NodeWithQuery, _super);
        function NodeWithQuery() {
            var _this = this;
            _super.apply(this, arguments);
            this.addQuery = function (query) { return new AddQuery(_this.path, queryUnion(_this.query, query), _this.actions); };
        }
        return NodeWithQuery;
    }(Node));
    RequestBuilder.NodeWithQuery = NodeWithQuery;
    var Attachment = (function (_super) {
        __extends(Attachment, _super);
        function Attachment() {
            _super.apply(this, arguments);
        }
        return Attachment;
    }(NodeWithQuery));
    RequestBuilder.Attachment = Attachment;
    var Attachments = (function (_super) {
        __extends(Attachments, _super);
        function Attachments() {
            _super.apply(this, arguments);
        }
        return Attachments;
    }(NodeWithQuery));
    RequestBuilder.Attachments = Attachments;
    var NodeWithAttachments = (function (_super) {
        __extends(NodeWithAttachments, _super);
        function NodeWithAttachments() {
            var _this = this;
            _super.apply(this, arguments);
            this.attachment = function (attachmentId) { return new Attachment(_this.path + "/attachments/" + attachmentId); };
            this.attachments = new Attachments(this.path + "/attachments");
        }
        return NodeWithAttachments;
    }(NodeWithQuery));
    RequestBuilder.NodeWithAttachments = NodeWithAttachments;
    var Message = (function (_super) {
        __extends(Message, _super);
        function Message() {
            _super.apply(this, arguments);
            this.actions = {
                GET: new Action(this.pathWithQuery),
                POST: new Action(this.pathWithQuery)
            };
        }
        return Message;
    }(NodeWithAttachments));
    RequestBuilder.Message = Message;
    var Messages = (function (_super) {
        __extends(Messages, _super);
        function Messages() {
            _super.apply(this, arguments);
            this.actions = {
                GETCOLLECTION: new Action(this.pathWithQuery),
                POST: new Action(this.pathWithQuery)
            };
        }
        return Messages;
    }(NodeWithQuery));
    RequestBuilder.Messages = Messages;
    var Event = (function (_super) {
        __extends(Event, _super);
        function Event() {
            _super.apply(this, arguments);
        }
        return Event;
    }(NodeWithAttachments));
    RequestBuilder.Event = Event;
    var Events = (function (_super) {
        __extends(Events, _super);
        function Events() {
            _super.apply(this, arguments);
        }
        return Events;
    }(NodeWithQuery));
    RequestBuilder.Events = Events;
    var User = (function (_super) {
        __extends(User, _super);
        function User() {
            var _this = this;
            _super.apply(this, arguments);
            this.message = function (messageId) { return new Message(_this.path + "/messages/" + messageId); };
            this.messages = new Messages(this.path + "/messages");
            this.event = function (eventId) { return new Event(_this.path + "/events/" + eventId); };
            this.events = new Events(this.path + "/events");
            this.calendarView = function (startDate, endDate) { return new Events(_this.path + "/calendarView", "startDateTime=" + startDate.toISOString() + "&endDateTime=" + endDate.toISOString()); };
        }
        return User;
    }(NodeWithQuery));
    RequestBuilder.User = User;
    var Users = (function (_super) {
        __extends(Users, _super);
        function Users() {
            _super.apply(this, arguments);
        }
        return Users;
    }(NodeWithQuery));
    RequestBuilder.Users = Users;
    var Root = (function () {
        function Root() {
            this.me = new User("/me");
            this.user = function (userId) { return new User("/users/" + userId); };
            this.users = new Users("/users/");
        }
        return Root;
    }());
    RequestBuilder.Root = Root;
})(RequestBuilder || (RequestBuilder = {}));
var MockGraph = (function () {
    function MockGraph() {
    }
    MockGraph.prototype.getCollection = function (endpoint) {
        var action = endpoint.actions && endpoint.actions.GETCOLLECTION;
        if (!action) {
            console.log("no GETCOLLECTION endpoint, sorry!");
        }
        else {
            console.log("GETCOLLECTION path", action.pathWithQuery);
            return {};
        }
    };
    MockGraph.prototype.get = function (endpoint) {
        var action = endpoint.actions && endpoint.actions.GET;
        if (!action) {
            console.log("no GET endpoint, sorry!");
        }
        else {
            console.log("GET path", action.pathWithQuery);
            return {};
        }
    };
    MockGraph.prototype.post = function (endpoint, request) {
        var action = endpoint.actions && endpoint.actions.POST;
        if (!action) {
            console.log("no POST endpoint, sorry!");
        }
        else {
            console.log("POST path", action.pathWithQuery);
        }
    };
    MockGraph.prototype.patch = function (endpoint, request) {
        var action = endpoint.actions && endpoint.actions.PATCH;
        if (!action) {
            console.log("no PATCH endpoint, sorry!");
        }
        else {
            console.log("PATCH path", action.pathWithQuery);
        }
    };
    MockGraph.prototype.delete = function (endpoint) {
        var action = endpoint.actions.DELETE;
        if (!action) {
            console.log("no DELETE endpoint, sorry!");
        }
        else {
            console.log("DELETE path", action.pathWithQuery);
        }
    };
    return MockGraph;
}());
var rb = new RequestBuilder.Root();
var graph = new MockGraph();
/*
var foo = rb.me.message("123");
var bar = rb.me.messages.addQuery("foo")

graph.get(foo).messageField
graph.getCollection(bar);
graph.post(foo, {} as Kurve.MessageDataModel);
graph.delete(foo);
graph.patch(foo, {} as Kurve.MessageDataModel);
*/ 
