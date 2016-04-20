module kurve {
    export class Graph {
        public static foo = () => "test";
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

        get me() { return new User(this, this.baseUrl); }
        get users() { return new Users(this, this.baseUrl); }

        public Get<Model, N extends Node>(path:string, self:N, scopes?:string[]): Promise<Singleton<Model, N>, Error> {
            console.log("GET", path, scopes);
            var d = new Deferred<Singleton<Model, N>, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Singleton<Model, N>(jsonResult, self));
            });

            return d.promise;
         }

        public GetCollection<Model, N extends CollectionNode>(path:string, self:N, next:N, scopes?:string[]): Promise<Collection<Model, N>, Error> {
            console.log("GET collection", path, scopes);
            var d = new Deferred<Collection<Model, N>, Error>();

            this.get(path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Collection<Model,N>(jsonResult, self, next))
            });

            return d.promise;
         }

        public Post<Model, N extends Node>(object:Model, path:string, self:N, scopes?:string[]): Promise<Singleton<Model, N>, Error> {
            console.log("POST", path, scopes);
            var d = new Deferred<Singleton<Model, N>, Error>();
            
/*
            this.post(object, path, (error, result) => {
                var jsonResult = JSON.parse(result) ;

                if (jsonResult.error) {
                    var errorODATA = new Error();
                    errorODATA.other = jsonResult.error;
                    d.reject(errorODATA);
                    return;
                }

                d.resolve(new Response<Model, N>({}, self));
            });
*/
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

        public post(object:string, url: string, callback: PromiseCallback<string>, responseType?: string, scopes?:string[]): void {
/*
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = () => {
                if (xhr.readyState === 4)
                    if (xhr.status === 202)
                        callback(null, responseType ? xhr.response : xhr.responseText);
                    else
                        callback(this.generateError(xhr));
            }
            xhr.send(object)
            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(addTokenError);
                }
            }, scopes);
*/
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
} //remove during bundling

