// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

import { Deferred, Promise, PromiseCallback } from "./promises";
import { Identity, OAuthVersion, Error } from "./identity";
import { UserDataModel, ProfilePhotoDataModel, MessageDataModel, EventDataModel, GroupDataModel, MailFolderDataModel, AttachmentDataModel } from "./models"
import { Collection, UserNode, UsersNode, pathWithQuery } from "./requestbuilder";
/*
    export module Scopes {
        class Util {
            static rootUrl = "https://graph.microsoft.com/";
        }
        export class General {
            public static OpenId: string = "openid";
            public static OfflineAccess: string = "offline_access";
        }
        export class User {
            public static Read: string = Util.rootUrl + "User.Read";
            public static ReadWrite: string = Util.rootUrl + "User.ReadWrite";
            public static ReadBasicAll: string = Util.rootUrl + "User.ReadBasic.All";
            public static ReadAll: string = Util.rootUrl + "User.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "User.ReadWrite.All";
        }
        export class Contacts {
            public static Read: string = Util.rootUrl + "Contacts.Read";
            public static ReadWrite: string = Util.rootUrl + "Contacts.ReadWrite";
        }
        export class Directory {
            public static ReadAll: string = Util.rootUrl + "Directory.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "Directory.ReadWrite.All";
            public static AccessAsUserAll: string = Util.rootUrl + "Directory.AccessAsUser.All";
        }
        export class Group {
            public static ReadAll: string = Util.rootUrl + "Group.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "Group.ReadWrite.All";
            public static AccessAsUserAll: string = Util.rootUrl + "Directory.AccessAsUser.All";
        }
        export class Mail {
            public static Read: string = Util.rootUrl + "Mail.Read";
            public static ReadWrite: string = Util.rootUrl + "Mail.ReadWrite";
            public static Send: string = Util.rootUrl + "Mail.Send";
        }
        export class Calendars {
            public static Read: string = Util.rootUrl + "Calendars.Read";
            public static ReadWrite: string = Util.rootUrl + "Calendars.ReadWrite";
        }
        export class Files {
            public static Read: string = Util.rootUrl + "Files.Read";
            public static ReadAll: string = Util.rootUrl + "Files.Read.All";
            public static ReadWrite: string = Util.rootUrl + "Files.ReadWrite";
            public static ReadWriteAppFolder: string = Util.rootUrl + "Files.ReadWrite.AppFolder";
            public static ReadWriteSelected: string = Util.rootUrl + "Files.ReadWrite.Selected";
        }
        export class Tasks {
            public static ReadWrite: string = Util.rootUrl + "Tasks.ReadWrite";
        }
        export class People {
            public static Read: string = Util.rootUrl + "People.Read";
            public static ReadWrite: string = Util.rootUrl + "People.ReadWrite";
        }
        export class Notes {
            public static Create: string = Util.rootUrl + "Notes.Create";
            public static ReadWriteCreatedByApp: string = Util.rootUrl + "Notes.ReadWrite.CreatedByApp";
            public static Read: string = Util.rootUrl + "Notes.Read";
            public static ReadAll: string = Util.rootUrl + "Notes.Read.All";
            public static ReadWriteAll: string = Util.rootUrl + "Notes.ReadWrite.All";
        }
    }
*/
/*
    export interface NextLink<T> {
        (callback? : PromiseCallback<T>): Promise<T, Error>;
    }

	export enum AttachmentType {
		fileAttachment,
		itemAttachment,
		referenceAttachment
	}

    export class Attachment extends DataModelWrapper<AttachmentDataModel>{
        public getType() : AttachmentType {
            switch (this._data['@odata.type']) {
                case "#microsoft.graph.fileAttachment":
                    return AttachmentType.fileAttachment;
                case "#microsoft.graph.itemAttachment":
                    return AttachmentType.itemAttachment;
                case "#microsoft.graph.referenceAttachment":
                    return AttachmentType.referenceAttachment;
            }
        }
    }
*/
    export class Graph {
        private req: XMLHttpRequest = null;
        private accessToken: string = null;
        private KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/v1.0";

        constructor(identityInfo: { identity: Identity });
        constructor(identityInfo: { defaultAccessToken: string });
        constructor(identityInfo: any) {
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }

        GET = <Model>(path:string, scopes?:string[]) => () => this.Get<Model>(path, scopes);
        GETCOLLECTION = <Model>(path:string, scopes?:string[]) => () => this.GetCollection<Model>(path, scopes);

        me = new UserNode(this, this.baseUrl);
        user = (userId:string) => new UserNode(this, this.baseUrl, userId);
        users = new UsersNode(this, this.baseUrl);

        public Get<Model>(path:string, scopes?:string[]): Promise<Model, Error> {
            console.log("GET", path);
            var d = new Deferred<Model, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(jsonResult);
            });

            return d.promise;
         }

        public GetCollection<Model>(path:string, scopes?:string[]): Promise<Collection<Model>, Error> {
            console.log("GETCOLLECTION", path);
            var d = new Deferred<Collection<Model>, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                var resultsArray = (jsonResult.value ? jsonResult.value : [jsonResult]) as any[];

                d.resolve({objects:resultsArray});
            });

            return d.promise;
         }
 
        //Only adds scopes when linked to a v2 Oauth of kurve identity
        private scopesForV2(scopes: string[]): string[] {
            if (!this.KurveIdentity)
                return null;
            if (this.KurveIdentity.getCurrentOauthVersion() === OAuthVersion.v1)
                return null;
            else return scopes;
        }

        //http verbs
        public getAsync(url: string): Promise<string, Error> {
            var d = new Deferred<string,Error>();
            this.get(url, (error, response) => error ? d.reject(error) : d.resolve(response))
            return d.promise;
        }

        public get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?:string[]): void {
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4)
                    if (xhr.status === 200)
                        callback(null, responseType ? xhr.response : xhr.responseText);
                    else
                        callback(this.generateError(xhr));
            }

            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(addTokenError);
                }
            }, scopes);
        }

        private generateError(xhr: XMLHttpRequest): Error {
            var response = new Error();
            response.status = xhr.status;
            response.statusText = xhr.statusText;
            if (xhr.responseType === '' || xhr.responseType === 'text')
                response.text = xhr.responseText;
            else
                response.other = xhr.response;
            return response;

        }
/*
        //Private methods

        private getUsers(urlString, callback: PromiseCallback<Users>, scopes?: string[], basicProfileOnly = true): void {
            this.get(urlString, (errorGet: Error, result: string) => {
                if (errorGet) {
                    callback(errorGet);
                    return;
                }

                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = usersODATA.error;
                    callback(errorODATA);
                    return;
                }

                var resultsArray = (usersODATA.value ? usersODATA.value : [usersODATA]) as any[];
                var users = new Users(this, resultsArray.map(o => new User(this, o)));
                var nextLink = usersODATA['@odata.nextLink'];
                if (nextLink) {
                    users.nextLink = (callback?: PromiseCallback<Users>) => {
                        var scopes = basicProfileOnly ? [Scopes.User.ReadBasicAll] : [Scopes.User.ReadAll];
                        var d = new Deferred<Users,Error>();
                        this.getUsers(nextLink, (error: Error, result: Users) => {
                            if (callback)
                                callback(error, result);
                            else
                                error ? d.reject(error) : d.resolve(result);
                        }, this.scopesForV2(scopes), basicProfileOnly);
                        return d.promise;
                    }
                }

                callback(null, users);
            },null,scopes);
        }

*/
        private addAccessTokenAndSend(xhr: XMLHttpRequest, callback: (error: Error) => void, scopes?:string[]): void {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            } else {
                //Using the integrated Identity object

                if (scopes) {
                    //v2 scope based tokens
                    this.KurveIdentity.getAccessTokenForScopes(scopes,false, ((token: string, error: Error) => {
                        if (error)
                            callback(error);
                        else {
                            xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                            xhr.send();
                            callback(null);
                        }
                    }));

                }
                else {
                    //v1 resource based tokens
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, ((error: Error, token: string) => {
                        if (error)
                            callback(error);
                        else {
                            xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                            xhr.send();
                            callback(null);
                        }
                    }));
                }
            }
        }

    }

//*********************************************************
//
//Kurve js, https://github.com/microsoftdx/kurvejs
//
//Copyright (c) Microsoft Corporation
//All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:




// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.




// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
//*********************************************************


/*
var graph = new Graph(new Identity({}));

graph.me.messages.GetMessages()
.then(collection => collection.objects[0].subject)
*/