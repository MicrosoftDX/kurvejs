import {Component} from 'angular2/core';
import {KurveService} from './kurve.service';

@Component({
    selector: 'my-app',
    template: `
    <h1>KurveJS Angular 2 example</h1>
        <button (click)="doLogin()">Initialize</button>
  `,
    providers: [KurveService]
})


export class AppComponent implements OnInit {
  
    constructor(private _kurveService: KurveService) { }
    doLogin() {
        this._kurveService.doLogin().then((result) => {
            window.alert("Hello user. Here's your id token:" + JSON.stringify(this._kurveService.getIdToken()));
        });
    }
}