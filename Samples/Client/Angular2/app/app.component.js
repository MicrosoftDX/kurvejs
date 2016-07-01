"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var core_1 = require('@angular/core');
var kurve = require('kurvejs');
var AppComponent = (function () {
    function AppComponent() {
        this.isAuthenticated = false;
    }
    //v2 Login
    AppComponent.prototype.login = function () {
        var _this = this;
        var id = new kurve.Identity("ed9162d5-0868-434b-a019-80fdb1add90c", "http://localhost:3000/node_modules/kurvejs/dist/login.html", { endpointVersion: kurve.EndpointVersion.v2 });
        id.loginAsync([kurve.Scopes.Mail.Read]).then(function (_) {
            _this.isAuthenticated = true;
            var graph = new kurve.Graph(id);
            graph.me.messages.GetMessages().then(function (data) {
                _this.messages = data.value;
            });
        });
    };
    AppComponent = __decorate([
        core_1.Component({
            selector: 'my-app',
            templateUrl: 'app/view-main.html'
        }), 
        __metadata('design:paramtypes', [])
    ], AppComponent);
    return AppComponent;
}());
exports.AppComponent = AppComponent;
//# sourceMappingURL=app.component.js.map