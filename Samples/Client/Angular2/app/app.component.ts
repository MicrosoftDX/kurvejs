import {Component} from '@angular/core';
import kurve = require('kurvejs');

@Component({
    selector: 'my-app',
    templateUrl: 'app/view-main.html'
})
export class AppComponent { 
    private isAuthenticated:boolean = false;
    private messages:kurve.MessageDataModel[];

    //v2 Login
    login() {
        const id = new kurve.Identity("ed9162d5-0868-434b-a019-80fdb1add90c", "http://localhost:3000/node_modules/kurvejs/dist/login.html", 
        {endpointVersion: kurve.EndpointVersion.v2});

        id.loginAsync([kurve.Scopes.Mail.Read]).then(_ => {
            this.isAuthenticated = true;

            const graph = new kurve.Graph(id);
            graph.me.messages.GetMessages().then(data => {
                this.messages = data.value
            });
        });
    }

    /*
    //v1 Login
    login() {
        const id = new kurve.Identity("1788837a-8e69-4d00-8fb1-a88e0f1ec4b9", "http://localhost:3000/node_modules/kurvejs/dist/login.html", 
        {endpointVersion: kurve.EndpointVersion.v1});

        id.loginAsync().then(_ => {
            this.isAuthenticated = true;

            const graph = new kurve.Graph(id);
            graph.me.messages.GetMessages().then(data => {
                this.messages = data.value
            });
        });
    }
    */
}
