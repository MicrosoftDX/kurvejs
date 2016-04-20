module kurve {

export interface ItemBody {
    contentType?: string;
    content?: string;
}

export interface EmailAddress {
    name?: string;
    address?: string;
}

export interface Recipient {
    emailAddress?: EmailAddress;
}

export interface UserDataModel {
    businessPhones?: string;
    displayName?: string;
    givenName?: string;
    jobTitle?: string;
    mail?: string;
    mobilePhone?: string;
    officeLocation?: string;
    preferredLanguage?: string;
    surname?: string;
    userPrincipalName?: string;
    id?: string;
}

export interface ProfilePhotoDataModel {
    id?: string;
    height?: Number;
    width?: Number;
}

export interface MessageDataModel {
    attachments?: AttachmentDataModel[];
    bccRecipients?: Recipient[];
    body?: ItemBody;
    bodyPreview?: string;
    categories?: string[]
    ccRecipients?: Recipient[];
    changeKey?: string;
    conversationId?: string;
    createdDateTime?: string;
    from?: Recipient;
    graph?: any;
    hasAttachments?: boolean;
    id?: string;
    importance?: string;
    isDeliveryReceiptRequested?: boolean;
    isDraft?: boolean;
    isRead?: boolean;
    isReadReceiptRequested?: boolean;
    lastModifiedDateTime?: string;
    parentFolderId?: string;
    receivedDateTime?: string;
    replyTo?: Recipient[]
    sender?: Recipient;
    sentDateTime?: string;
    subject?: string;
    toRecipients?: Recipient[];
    webLink?: string;
}

export interface Attendee {
    status?: ResponseStatus;
    type?: string;
    emailAddress?: EmailAddress;
}

export interface DateTimeTimeZone {
    dateTime?: string;
    timeZone?: string;
}

export interface PatternedRecurrence { }

export interface ResponseStatus {
    response?: string;
    time?: string
}

export interface Location {
    displayName?: string;
    address?: any;
}

export interface EventDataModel {
    attendees?: Attendee[];
    body?: ItemBody;
    bodyPreview?: string;
    categories?: string[];
    changeKey?: string;
    createdDateTime?: string;
    end?: DateTimeTimeZone;
    hasAttachments?: boolean;
    iCalUId?: string;
    id?: string;
    IDBCursor?: string;
    importance?: string;
    isAllDay?: boolean;
    isCancelled?: boolean;
    isOrganizer?: boolean;
    isReminderOn?: boolean;
    lastModifiedDateTime?: string;
    location?: Location;
    organizer?: Recipient;
    originalEndTimeZone?: string;
    originalStartTimeZone?: string;
    recurrence?: PatternedRecurrence;
    reminderMinutesBeforeStart?: number;
    responseRequested?: boolean;
    responseStatus?: ResponseStatus;
    sensitivity?: string;
    seriesMasterId?: string;
    showAs?: string;
    start?: DateTimeTimeZone;
    subject?: string;
    type?: string;
    webLink?: string;
}

export interface GroupDataModel {
    id?: string;
    description?: string;
    displayName?: string;
    groupTypes?: string[];
    mail?: string;
    mailEnabled?: Boolean;
    mailNickname?: string;
    onPremisesLastSyncDateTime?: Date;
    onPremisesSecurityIdentifier?: string;
    onPremisesSyncEnabled?: Boolean;
    proxyAddresses?: string[];
    securityEnabled?: Boolean;
    visibility?: string;
}

export interface MailFolderDataModel {
    id?: string;
    displayName?: string;
    childFolderCount?: number;
    unreadItemCount?: number;
    totalItemCount?: number;
}

export interface AttachmentDataModel {
    contentId?: string;
    id?: string;
    isInline?: boolean;
    lastModifiedDateTime?: Date;
    name?: string;
    size?: number;

    /* File Attachments */
    contentBytes?: string;
    contentLocation?: string;
    contentType?: string;
}

} //remove during bundling