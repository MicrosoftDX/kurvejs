// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {
     export class Identity {
        public authContext: any = null;
        public config: any = null;
        public isCallback: boolean = false;
        public clientId: string;
        public tenantId: string;
        private req: XMLHttpRequest;
        private state: string;
        private nonce: string;
        private idToken: any;
        private loginCallback: (error: string) => void;
        private getTokenCallback: (token: string, error: string) => void;
        private redirectUri: string;
        private tokenCache: any;
        private logonUser: any;
        private refreshTimer: any;
    
        constructor(tenantId = "", clientId = "", redirectUri = "") {
            this.tenantId = tenantId;
            this.clientId = clientId;
            this.redirectUri = redirectUri;
            this.req = new XMLHttpRequest();
            this.tokenCache = {};

            //Callback handler from other windows
            window.addEventListener("message", ((event: MessageEvent) => {
                if (event.data.type === "id_token") {
                    //Callback being called by the login window
                    if (!event.data.token) {
                        this.loginCallback(event.data);
                    }
                    else {
                        //check for state
                        if (this.state !== event.data.state) {
                            throw "Invalid state";
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
                        var token:string = event.data.token;
                        var iframe = document.getElementById("tokenIFrame");
                        iframe.parentNode.removeChild(iframe);

                        if (event.data.state !== this.state) {
                            throw "Invalid state";
                        }
                        else {
                            this.getTokenCallback(token, null);
                        }
                    }
                }
            }));
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
            this.refreshTimer = setTimeout((() => {
                this.renewIdToken();
            }), parseInt(decodedTokenJSON.exp) * 1000 - 300000); 
        }

        private decodeAccessToken(accessToken: string, resource:string): void {
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

        public getAccessTokenAsync(resource: string): Promise {

            var d = new Deferred();
            this.getAccessToken(resource, ((token, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        }

        public getAccessToken(resource: string, callback: (token: string, error: string) => void): void {
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
            this.getTokenCallback = ((token: string, error: string) => {
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
            iframe.src = "./login.html?tenantId=" + encodeURIComponent(this.tenantId) +
            "&clientId=" + encodeURIComponent(this.clientId) +
            "&resource=" + encodeURIComponent(resource) +
            "&redirectUri=" + encodeURIComponent(this.redirectUri) +
            "&state=" + encodeURIComponent(this.state) +
            "&nonce=" + encodeURIComponent(this.nonce);
            document.body.appendChild(iframe);
        }

        public loginAsync(): Promise {
            var d = new Deferred();
            this.login(((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve();
                }
            }));
            return d.promise;
        }

        public login(callback: (error: string) => void): void {
            this.loginCallback = callback;
            this.state = this.generateNonce();
            this.nonce = this.generateNonce();
            window.open("./login.html?tenantId=" + encodeURIComponent(this.tenantId) +
                "&clientId=" + encodeURIComponent(this.clientId) +
                "&redirectUri=" + encodeURIComponent(this.redirectUri) +
                "&state=" + encodeURIComponent(this.state) +
                "&nonce=" + encodeURIComponent(this.nonce), "_blank");
        }

        public logOut(): void {
            var url = "https://login.windows.net/" + this.tenantId + "/oauth2/logout?post_logout_redirect_uri=" + encodeURI(window.location.href);
            window.location.href = url;
        }

        private base64Decode(encodedString: string): string {
            var e = {}, i, b = 0, c, x, l = 0, a, r = '', w = String.fromCharCode, L = encodedString.length;
            var A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            for (i = 0; i < 64; i++) { e[A.charAt(i)] = i; }
            for (x = 0; x < L; x++) {
                c = e[encodedString.charAt(x)]; b = (b << 6) + c; l += 6;
                while (l >= 8) { ((a = (b >>> (l -= 8)) & 0xff) || (x < (L - 2))) && (r += w(a)); }
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
//Kurve js, https://github.com/matvelloso/kurve-js
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
