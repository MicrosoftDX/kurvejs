﻿import {Injectable} from 'angular2/core';
@Injectable()
export class KurveService {
    private identity: kurve.Identity;
    doLogin() {
        var id = "6f95d630-018e-441d-8d19-fa78a1232b42";
        var redirectURI = "https://localhost:44300/login.html";

        this.identity = new kurve.Identity({ clientId: id, tokenProcessingUri: redirectURI, version: kurve.EndPointVersion.v2 });

        return this.identity.loginAsync();
    };
 
    getIdToken() {
 

        return this.identity.getIdToken();
    }
 
}