namespace Kurve {

    export class Graph {
        private req: XMLHttpRequest = null;
        private accessToken: string;
        private KurveIdentity: Identity;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private https: any;
  
        root: string = "https://graph.microsoft.com/v1.0";
        mode: Mode = Mode.Client;
        endpointVersion: EndpointVersion = EndpointVersion.v1;

        constructor(id: Identity, options?: {root?: string})
        constructor(id: string, options?: {root?: string, endpointVersion: EndpointVersion, mode?: Mode, https?: any, })
        constructor(id: any, options:any) {
            if (typeof(id) === "string") {
                this.accessToken = id;
                if (options && options.mode == Mode.Node) {
                    this.mode = Mode.Node;
                    if (options && options.https)
                        this.https = options.https;
                }
            } else {
                this.KurveIdentity = id;
                this.mode = id.mode;
                this.endpointVersion = id.endpointVersion;
                this.https = id.https;
            }
            if (options && options.root)
                this.root = options.root;
            console.log("graph",this);
        }

        get me() { return new User(this); }
        get users() { return new Users(this); }
        get groups() { return new Groups(this); }

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
                           'Content-Type': responseType?responseType:'application/json',
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

}

