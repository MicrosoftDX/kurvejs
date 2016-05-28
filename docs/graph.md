# Kurve QueryBuilder

Consider the state of developing against the Microsoft Graph today. For every operation you wish to perform you must look in the documentation to find:

1. The path, e.g. "me/manager"
2. The access requirements, e.g. "User.Read.All"
3. The formats of the request and response objects

It's another trip through the documentation to find out other operations that can be performed at the same endpoint, or on the response object.

This is what tooling is supposed to help with. This is the problem that Kurve's QueryBuilder solves: leverage Intellisense to let you to easily discover and access the Microsoft Graph.

Here's what it looks like. With a Kurve _graph_ object in hand, just start typing in a TypeScript-aware editor such as Visual Studio Code:

    You type...                 The editor prompts you with...
    
    graph.                      me, users
    graph.me.                   events, messages, calendarView, mailFolders, GetUser
    graph.me.events.            GetEvents, $
    graph.me.events.$           (eventId:string) => Event
    graph.me.events.$("123").   GetEvent

Nodes in the path (e.g. _events_) are lowerCamelCase. "$" is where you type an id. Operations (e.g. _GetEvent_) are UpperCamelCase.    

Each endpoint exposes the set of available Graph operations through strongly typed methods:

    graph.me.GetUser() => Kurve.UserDataModel
        GET "/me"
    graph.me.events.GetEvents() => Kurve.EventDataModel[]
        GET "/me/events"
    graph.me.events.CreateEvent(event:Kurve.EventDataModel)
        POST "/me/events"

Certain Graph endpoints are implemented as OData "Functions". These are not treated as Graph nodes. They're just methods: 

    graph.me.photo.GetPhotoImage()
        POST "/me/photo/$value"

Graph operations are exposed through Promises:

    graph.me.messages.GetMessages().then(messages =>
        messages.forEach(message =>
            console.log(message.subject)
        )
    )

## Navigating the Graph path from responses

Normal usage of the MS Graph requires building a new path after each operation. If you wish to enumerate through message, outputting the subject line and then the sizes of all attachments, you'd do:

    GET "me/messages/${messageId}"
    GET "me/messages/${messageId}/attachments"
    GET "me/messages/${messageId}/attachments/${attachmentId}"

QueryBuilder helps you write more concise code. All response objects return a "_context" property which represent the Graph path location of the returned object, allowing you to pick up where you left off:

    graph.me.messages.$("123").GetMessage().then(message =>
        console.log(message.subject);
        message._context.attachments.GetAttachments().then(attachments => // message._context === graph.me.messages.$("123")
            attachments.forEach(attachment => 
                console.log(attachment.contentBytes)
            )
        )
    )

Members of returned collections also have a "_context" property:

    graph.me.messages.GetMessages().then(messages =>
        messages.forEach(message =>
            message._context.attachments.GetAttachments().then(attachments => // message._context === graph.me.messages.$(message.id)
                attachments.forEach(attachment =>
                    console.log(attachment.id);
                )
            )
        )
    )

Now consider these two cases:

    GET "me/manager"
    GET "me/directReports/{userId}"
    
In both cases, GetUser() returns a user object, but you can't do user operations on that place in the graph.
For example, these operations are invalid:

    GET "me/manager/directReports"
    GET "me/directReports/{userId}/manager"
    
Instead you must start again at the beginning:

    GET "users/{userId}/directReports"
    GET "users/{userId}/manager"
    
We facilitate this by setting the "_context" property intelligently: 

    graph.me.manager.GetUser().then(user =>
        user._context.directReports.GetUsers().then(users =>   // user._context === users.$(user.id)
            users.forEach(user =>
                console.log(user.displayName)
            )
        )
    )

    graph.me.directReports.GetUsers().then(users =>
        users.forEach(user =>
            user._context.manager.GetManager().then(manager =>     // user._context === users.$(user.id)
                console.log(manager.displayName)
            )
        )
    )

In other words, QueryBuilder lets you continue to do operations on the returned object. 

## Paginated Collections

Operations which return paginated collections can return a "_next" request object. This can be utilized in a recursive function:

    ListMessageSubjects(messages:Kurve.GraphCollection<Kurve.MessageDataModel, Kurve.Messages, Kurve.Message>) {
        messages.forEach(message => console.log(message.subject));
        if (messages._next)
            messages._next().then(nextMessages =>
                ListMessageSubjects(nextMessages)
            )
        })
    }
    graph.me.messages.GetMessages().then(messages =>
        ListMessageSubjects(messages)
    );
    
(With async/await support, an iteration pattern can be used intead of recursion)

Collections also expose a "_raw" property which allows access to the actual JSON returned from the Microsoft Graph.

    graph.users.GetUsers().then(users =>
        console.log(users._raw["@odata.context"])
    ); 

## About Response Types

The response types seem complex, but they are actually very simple. For example, in:

    graph.me.GetUser().then(user =>

_user_ is type _GraphObject&lt;UserDataModel, User>_ which is just an alias for _UserDataModel & { _context: User }_. In other words, it's a standard Microsoft Graph User object, plus one additional field called __context_, described above.

Similarly for collections:

    graph.users.Getusers.then(users =>

_users_ is type _GraphCollection&lt;UserDataModel, Users, User>_ which is just an alias for _UserDataModel[] & { _context: User, _next?: () => GraphCollection&lt;UserDataModel, Users, User>, _raw:any }_. In other words, it's an array of Microsoft Graph User objects, plus three extra fields described above.

Because _GraphObject_ and _GraphCollection_ are defined as type intersections, they can be used anywhere the base types are expected, e.g.

    const showDisplayName = (user:UserDataModel) => console.log(user.displayName);
    graph.me.GetUser().then(user => showDisplayName(user));
    
    const showDisplayNames = (users:UserDataModel[]) => users.forEach(user => console.log(user.displayName));
    graph.users.GetUsers().then(users => showDisplayNames(users));

## OData

Every Graph operation may include OData queries:

    graph.me.messages.GetMessages("$select=subject,id&$orderby=id")
        GET "/me/messages/$select=subject,id&$orderby=id"

There is an optional OData helper to aid in constructing more complex queries:

    graph.me.messages.GetMessages(new OData()
        .select("subject", "id")
        .orderby("id")
    )
        GET "/me/messages/$select=subject,id&$orderby=id"

Some nodes expose node-specific ODATA helpers: 

    graph.me.calendarView.GetCalendarView(CalendarView.dateRange(start, end))
        GET "/me/calendarView?startDateTime={start}&endDateTime={end}"
