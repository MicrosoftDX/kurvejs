export interface ItemBody {
    contentType: string;
    content: string;
}

export interface EmailAddress {
    name: string;
    address: string;
}

export interface Recipient {
    emailAddress: EmailAddress;
}

export class UserDataModel {
    public businessPhones: string;
    public displayName: string;
    public givenName: string;
    public jobTitle: string;
    public mail: string;
    public mobilePhone: string;
    public officeLocation: string;
    public preferredLanguage: string;
    public surname: string;
    public userPrincipalName: string;
    public id: string;
}

export class ProfilePhotoDataModel {
    public id: string;
    public height: Number;
    public width: Number;
}

export class MessageDataModel {
    attachments: AttachmentDataModel[];
    bccRecipients: Recipient[];
    body: ItemBody;
    bodyPreview: string;
    categories: string[]
    ccRecipients: Recipient[];
    changeKey: string;
    conversationId: string;
    createdDateTime: string;
    from: Recipient;
    graph: any;
    hasAttachments: boolean;
    id: string;
    importance: string;
    isDeliveryReceiptRequested: boolean;
    isDraft: boolean;
    isRead: boolean;
    isReadReceiptRequested: boolean;
    lastModifiedDateTime: string;
    parentFolderId: string;
    receivedDateTime: string;
    replyTo: Recipient[]
    sender: Recipient;
    sentDateTime: string;
    subject: string;
    toRecipients: Recipient[];
    webLink: string;
}

export interface Attendee {
    status: ResponseStatus;
    type: string;
    emailAddress: EmailAddress;
}

export interface DateTimeTimeZone {
    dateTime: string;
    timeZone: string;
}

export interface PatternedRecurrence { }

export interface ResponseStatus {
    response: string;
    time: string
}

export interface Location {
    displayName: string;
    address: any;
}

export class EventDataModel {
    attendees: Attendee[];
    body: ItemBody;
    bodyPreview: string;
    categories: string[];
    changeKey: string;
    createdDateTime: string;
    end: DateTimeTimeZone;
    hasAttachments: boolean;
    iCalUId: string;
    id: string;
    IDBCursor: string;
    importance: string;
    isAllDay: boolean;
    isCancelled: boolean;
    isOrganizer: boolean;
    isReminderOn: boolean;
    lastModifiedDateTime: string;
    location: Location;
    organizer: Recipient;
    originalEndTimeZone: string;
    originalStartTimeZone: string;
    recurrence: PatternedRecurrence;
    reminderMinutesBeforeStart: number;
    responseRequested: boolean;
    responseStatus: ResponseStatus;
    sensitivity: string;
    seriesMasterId: string;
    showAs: string;
    start: DateTimeTimeZone;
    subject: string;
    type: string;
    webLink: string;
}

export class GroupDataModel {
    public id: string;
    public description: string;
    public displayName: string;
    public groupTypes: string[];
    public mail: string;
    public mailEnabled: Boolean;
    public mailNickname: string;
    public onPremisesLastSyncDateTime: Date;
    public onPremisesSecurityIdentifier: string;
    public onPremisesSyncEnabled: Boolean;
    public proxyAddresses: string[];
    public securityEnabled: Boolean;
    public visibility: string;
}

export class MailFolderDataModel {
    public id: string;
    public displayName: string;
    public childFolderCount: number;
    public unreadItemCount: number;
    public totalItemCount: number;
}


export class AttachmentDataModel {
    public contentId: string;
    public id: string;
    public isInline: boolean;
    public lastModifiedDateTime: Date;
    public name: string;
    public size: number;

    /* File Attachments */
    public contentBytes: string;
    public contentLocation: string;
    public contentType: string;
}
