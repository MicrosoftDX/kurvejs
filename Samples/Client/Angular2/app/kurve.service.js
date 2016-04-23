"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var core_1 = require('angular2/core');
var KurveService = (function () {
    function KurveService() {
    }
    KurveService.prototype.doLogin = function () {
        var id = "6f95d630-018e-441d-8d19-fa78a1232b42";
        var redirectURI = "https://localhost:44300/login.html";
        this.identity = new kurve.Identity({ clientId: id, tokenProcessingUri: redirectURI, version: kurve.EndPointVersion.v2 });
        return this.identity.loginAsync();
    };
    ;
    KurveService.prototype.getIdToken = function () {
        return this.identity.getIdToken();
    };
    KurveService = __decorate([
        core_1.Injectable()
    ], KurveService);
    return KurveService;
}());
exports.KurveService = KurveService;
