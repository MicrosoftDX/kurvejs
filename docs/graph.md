Kurve's QueryBuilder allows you to easily discover and access the Microsoft Graph.

Just start typing in Visual Studio Code and other TypeScript-aware editors too see how intellisense helps you explore the graph:

    const graph = new Graph(...); 

    graph.                      me, users
    graph.me.                   events, messages, calendarView, mailFolders, GetUser
    graph.me.events.            GetEvents, $
    graph.me.events.$           (eventId:string) => Event
    graph.me.events.$("123").   GetEvent

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

All operations return a "_context" property which represent the current Graph path location:

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
        user._context.directReports.GetUsers().then(users =>   // user._context === users.${user.id}
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

The returned types are necessarily somewhat complex. You can access the simple models through _item(s)

    graph.users.${userid}.GetUser().then(user =>
        // user._item is type UserDataModel
    )
    graph.users.GetUsers().then(users =>
        // users._items is type Kurve.UserDataModel[]
    )

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
