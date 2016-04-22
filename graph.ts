module kurve {
    export class Graph {
        private req: XMLHttpRequest = null;
        private accessToken: string = null;
        KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/v1.0";
        private https: any;
        private mode: Mode;

        constructor(identityInfo: { identity: Identity }, mode: Mode, https?: any);
        constructor(identityInfo: { defaultAccessToken: string },mode:Mode, https?: any);
        constructor(identityInfo: any, mode: Mode, https?: any) {
            if (https) this.https = https;
            this.mode = mode;
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }

        get me() { return new User(this, this.baseUrl); }
        get users() { return new Users(this, this.baseUrl); }
        get groups() { return new Groups(this, this.baseUrl); }

        public Get<Model, N extends Node>(path:string, self:N, scopes?:string[], responseType?:string): Promise<Singleton<Model, N>, Error> {
            console.log("GET", path, scopes);
            var d = new Deferred<Singleton<Model, N>, Error>();

            this.get(path, (error, result) => {
                if (!responseType){
                    var jsonResult = JSON.parse(result) ;

                    if (jsonResult.error) {
                        var errorODATA = new Error();
                        errorODATA.other = jsonResult.error;
                        d.reject(errorODATA);
                        return;
                    }
                    d.resolve(new Singleton<Model, N>(jsonResult, self));
                } else {
                    d.resolve(new Singleton<Model, N>(result, self));
                }

                
            }, responseType, scopes);

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
            }, null, scopes);

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
 
        public get(url: string, callback: PromiseCallback<string>, responseType?: string, scopes?: string[]): void {
            if (this.mode === Mode.Client) {
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
            } else {
                var token = this.findAccessToken((token, error) => {
                    var path = url.substr(27, url.length);
                    var options = {
                        host: 'graph.microsoft.com',
                        port: '443',
                        path: path,
                        method: 'GET',
                        headers: {
                            'Content-Type': responseType,
                            accept: '*/*',
                            'Authorization': 'Bearer ' + token
                        }
                    };

                    var post_req = this.https.request(options, (response) => {
                        response.setEncoding('utf8');
                        response.on('data', (chunk) => {
                            callback(null, chunk);
                        });
                    });

                    post_req.end();

                }, scopes);
            }
        }
        private findAccessToken(callback: (token: string, error: Error) => void, scopes?: string[]): void {
            if (this.accessToken) {
                callback(this.accessToken, null);
            } else {
                //Using the integrated Identity object

                if (scopes) {
                    //v2 scope based tokens
                    this.KurveIdentity.getAccessTokenForScopes(scopes, false, ((token: string, error: Error) => {
                        if (error)
                            callback(null, error);
                        else {
                            callback(token, null);
                        }
                    }));

                }
                else {
                    //v1 resource based tokens
                    this.KurveIdentity.getAccessToken(this.defaultResourceID, ((error: Error, token: string) => {
                        if (error)
                            callback(null, error);
                        else {
                            callback(token, null);
                        }
                    }));
                }
            }
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

        private addAccessTokenAndSend(xhr: XMLHttpRequest, callback: (error: Error) => void, scopes?: string[]): void {
            this.findAccessToken((token, error) => {
                if (token) {
                    xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                    xhr.send();
                } else {
                    callback(error);
                }
            }, scopes);
        }

    }
} //remove during bundling

