"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var core_1 = require('angular2/core');
var kurve_service_1 = require('./kurve.service');
var AppComponent = (function () {
    function AppComponent(_kurveService) {
        this._kurveService = _kurveService;
    }
    AppComponent.prototype.doLogin = function () {
        var _this = this;
        this._kurveService.doLogin().then(function (result) {
            window.alert("Hello user. Here's your id token:" + JSON.stringify(_this._kurveService.getIdToken()));
        });
    };
    AppComponent = __decorate([
        core_1.Component({
            selector: 'my-app',
            template: "\n    <h1>KurveJS Angular 2 example</h1>\n        <button (click)=\"doLogin()\">Initialize</button>\n  ",
            providers: [kurve_service_1.KurveService]
        })
    ], AppComponent);
    return AppComponent;
}());
exports.AppComponent = AppComponent;
