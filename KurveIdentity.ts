// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {
    export class Error {
        public status: number;
        public statusText: string;
        public text: string;
        public other: any;
    }


    export class Identity {
        public authContext: any = null;
        public config: any = null;
        public isCallback: boolean = false;
        public clientId: string;
        private req: XMLHttpRequest;
        private state: string;
        private nonce: string;
        private idToken: any;
        private loginCallback: (error: Error) => void;
        private accessTokenCallback: (token: string, error: Error) => void;
        private getTokenCallback: (token: string, error: Error) => void;
        private tokenProcessingUri: string;
        private tokenCache: any;
        private logonUser: any;
        private refreshTimer: any;

        constructor(clientId = "", tokenProcessingUri = "") {
            this.clientId = clientId;
            this.tokenProcessingUri = tokenProcessingUri;
            this.req = new XMLHttpRequest();
            this.tokenCache = {};

            //Callback handler from other windows
            window.addEventListener("message", ((event: MessageEvent) => {
                if (event.data.type === "id_token") {
                    //Callback being called by the login window
                    if (!event.data.token) {
                        this.loginCallback(event.data); //BUG?  This is an error
                    }
                    else {
                        //check for state
                        if (this.state !== event.data.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            this.loginCallback(error);
                        } else {
                            this.decodeIdToken(event.data.token);
                            this.loginCallback(null);
                        }
                    }
                } else if (event.data.type === "access_token") {
                    //Callback being called by the iframe with the token
                    if (!event.data.token)
                        this.getTokenCallback(null, event.data);
                    else {
                        var token: string = event.data.token;
                        var iframe = document.getElementById("tokenIFrame");
                        iframe.parentNode.removeChild(iframe);

                        if (event.data.state !== this.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            this.getTokenCallback(null, error);
                        }
                        else {
                            this.getTokenCallback(token, null);
                        }
                    }
                }
            }));
        }


        public checkForIdentityRedirect(): boolean {
            function token(s: string) {
                var index = window.location.href.indexOf(s);
                return (index > 0) ? window.location.href.substring(index + s.length) : null;
            }

            function parseQueryString(str: string) {
                var queryString = str || window.location.search || '';
                var keyValPairs: any[] = [];
                var params: any = {};
                queryString = queryString.replace(/.*?\?/, "");

                if (queryString.length) {
                    keyValPairs = queryString.split('&');
                    for (var pairNum in keyValPairs) {
                        var key = keyValPairs[pairNum].split('=')[0];
                        if (!key.length) continue;
                        if (typeof params[key] === 'undefined')
                            params[key] = [];
                        params[key].push(keyValPairs[pairNum].split('=')[1]);
                    }
                }
                return params;
            }

            var params = parseQueryString(window.location.href);
            var idToken = token("#id_token=");
            var accessToken = token("#access_token");
            if (idToken) {
                if (true || this.state === params["state"][0]) { //BUG? When you are in a pure redirect system you don't remember your state or nonce so don't check.
                    this.decodeIdToken(idToken);
                    this.loginCallback(null);
                } else {
                    var error = new Error();
                    error.statusText = "Invalid state";
                    this.loginCallback(error);
                }
                return true;
            }
            else if (accessToken) {
                throw "Should not get here.  This should be handled via the iframe approach."
                if (this.state === params["state"][0]) {
                    this.getTokenCallback(accessToken, null);
                } else {
                    var error = new Error();
                    error.statusText = "Invalid state";
                    this.getTokenCallback(null, error);
                }
            }
            return false;
        }

        private decodeIdToken(idToken: string): void {

            var decodedToken = this.base64Decode(idToken.substring(idToken.indexOf('.') + 1, idToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            this.idToken = {
                token: idToken,
                expiry: expiryDate,
                upn: decodedTokenJSON.upn,
                tenantId: decodedTokenJSON.tid,
                family_name: decodedTokenJSON.family_name,
                given_name: decodedTokenJSON.given_name,
                name: decodedTokenJSON.name
            }
            var expiration: Number = expiryDate.getTime() - new Date().getTime() - 300000;

            this.refreshTimer = setTimeout((() => {
                this.renewIdToken();
            }), expiration);
        }

        private decodeAccessToken(accessToken: string, resource: string): void {
            var decodedToken = this.base64Decode(accessToken.substring(accessToken.indexOf('.') + 1, accessToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            this.tokenCache[resource] = {
                resource: resource,
                token: accessToken,
                expiry: expiryDate
            }
        }

        public getIdToken(): any {
            return this.idToken;
        }
        public isLoggedIn(): boolean {
            if (!this.idToken) return false;
            return (this.idToken.expiry > new Date());
        }

        private renewIdToken(): void {
            clearTimeout(this.refreshTimer);
            this.login((() => {

            }));
        }

        public getAccessTokenAsync(resource: string): Promise<string, Error> {

            var d = new Deferred<string, Error>();
            this.getAccessToken(resource, ((token, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        }



        public getAccessToken(resource: string, callback: (token: string, error: Error) => void): void {
            //Check for cache and see if we have a valid token
            var cachedToken = this.tokenCache[resource];
            if (cachedToken) {
                //We have it cached, has it expired? (5 minutes error margin)
                if (cachedToken.expiry > (new Date(new Date().getTime() + 60000))) {
                    callback(<string>cachedToken.token, null);
                    return;
                }
            }
            //If we got this far, we need to go get this token

            //Need to create the iFrame to invoke the acquire token
            this.getTokenCallback = ((token: string, error: Error) => {
                if (error) {
                    callback(null, error);
                }
                else {
                    this.decodeAccessToken(token, resource);
                    callback(token, null);

                }
            });

            this.nonce = this.generateNonce();
            this.state = this.generateNonce();

            var iframe = document.createElement('iframe');
            iframe.style.display = "none";
            iframe.id = "tokenIFrame";
            iframe.src = this.tokenProcessingUri +
                "?clientId=" + encodeURIComponent(this.clientId) +
                "&resource=" + encodeURIComponent(resource) +
                "&redirectUri=" + encodeURIComponent(this.tokenProcessingUri) +
                "&state=" + encodeURIComponent(this.state) +
                "&nonce=" + encodeURIComponent(this.nonce);
            document.body.appendChild(iframe);
        }

        public loginAsync(toUrl?: string): Promise<void, Error> {
            var d = new Deferred<void, Error>();
            this.login((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }, toUrl);
            return d.promise;
        }

        public login(callback: (error: Error) => void, toUrl?: string): void {
            this.loginCallback = callback;
            this.state = this.generateNonce();
            this.nonce = this.generateNonce();

            window.open(this.tokenProcessingUri +
                "?clientId=" + encodeURIComponent(this.clientId) +
                "&redirectUri=" + encodeURIComponent(this.tokenProcessingUri) +
                "&state=" + encodeURIComponent(this.state) +
                "&nonce=" + encodeURIComponent(this.nonce), "_blank");

        }


        public loginNoWindowAsync(): Promise<void, Error> {
            var d = new Deferred<void, Error>();
            this.loginNoWindow((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            });
            return d.promise;
        }

        public loginNoWindow(callback: (error: Error) => void): void {
            this.loginCallback = callback;
            this.state = this.generateNonce();
            this.nonce = this.generateNonce();

            var redirected = this.checkForIdentityRedirect();
            if (!redirected) {
                var toSelfUrl = window.location.href.split("#")[0];  // For no loginWindows come back here
                var url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token" +
                    "&client_id=" + encodeURIComponent(this.clientId) +
                    "&redirect_uri=" + encodeURIComponent(toSelfUrl) +
                    "&state=" + encodeURIComponent(this.state) +
                    "&nonce=" + encodeURIComponent(this.nonce);
                window.location.href = url;
            }
        }

        public logOut(): void {
            var url = "https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=" + encodeURI(window.location.href);
            window.location.href = url;
        }

        private base64Decode(encodedString: string): string {
            var e: any = {}, i: number, b = 0, c: number, x: number, l = 0, a: any, r = '', w = String.fromCharCode, L = encodedString.length;
            var A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            for (i = 0; i < 64; i++) { e[A.charAt(i)] = i; }
            for (x = 0; x < L; x++) {
                c = e[encodedString.charAt(x)];
                b = (b << 6) + c;
                l += 6;
                while (l >= 8) {
                    ((a = (b >>> (l -= 8)) & 0xff) || (x < (L - 2))) && (r += w(a));
                }
            }
            return r;
        }

        private generateNonce(): string {
            var text = "";
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            for (var i = 0; i < 32; i++) {
                text += chars.charAt(Math.floor(Math.random() * chars.length));
            }
            return text;
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
