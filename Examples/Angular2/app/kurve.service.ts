import {Injectable} from 'angular2/core';
@Injectable()
export class KurveService {
    
    doLogin() {
        var id = "6f95d630-018e-441d-8d19-fa78a1232b42";
        var redirectURI = "https://localhost:44300/login.html";

        var identity = new Kurve.Identity(id, redirectURI, Kurve.OAuthVersion.v2);

        return identity.loginAsync();
    }
 
}