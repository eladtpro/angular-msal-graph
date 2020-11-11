(window["webpackJsonp"] = window["webpackJsonp"] || []).push([["main"],{

/***/ 0:
/*!***************************!*\
  !*** multi ./src/main.ts ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__(/*! C:\Users\eladt\Documents\Projects\Angular\MSAL_Angular_Sample_Application\src\main.ts */"zUnb");


/***/ }),

/***/ "2hxB":
/*!********************************!*\
  !*** ./src/app/models/user.ts ***!
  \********************************/
/*! exports provided: User */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "User", function() { return User; });
var User = /** @class */ (function () {
    function User(user) {
        this.displayName = user.displayName;
        // Prefer the mail property, but fall back to userPrincipalName
        this.mail = user.mail || user.userPrincipalName;
        // this.timeZone = user.mailboxSettings.timeZone;
        // Use default avatar
        this.avatar = '/assets/no-profile-photo.png';
    }
    return User;
}());



/***/ }),

/***/ "AytR":
/*!*****************************************!*\
  !*** ./src/environments/environment.ts ***!
  \*****************************************/
/*! exports provided: environment */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "environment", function() { return environment; });
// This file can be replaced during build by using the `fileReplacements` array.
// `ng build --prod` replaces `environment.ts` with `environment.prod.ts`.
// The list of file replacements can be found in `angular.json`.
var environment = {
    production: false,
    graph: 'https://graph.microsoft.com/v1.0/me'
};
/*
 * For easier debugging in development mode, you can import the following file
 * to ignore zone related error stack frames such as `zone.run`, `zoneDelegate.invokeTask`.
 *
 * This import should be commented out in production mode because it will have a negative impact
 * on performance if an error is thrown.
 */
// import 'zone.js/dist/zone-error';  // Included with Angular CLI.
// eladt@b2cpm.onmicrosoft.com
// RequestManager1234


/***/ }),

/***/ "BuFo":
/*!***************************************************!*\
  !*** ./src/app/components/home/home.component.ts ***!
  \***************************************************/
/*! exports provided: HomeComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "HomeComponent", function() { return HomeComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");


var HomeComponent = /** @class */ (function () {
    function HomeComponent() {
    }
    HomeComponent.prototype.ngOnInit = function () {
    };
    HomeComponent.ɵfac = function HomeComponent_Factory(t) { return new (t || HomeComponent)(); };
    HomeComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: HomeComponent, selectors: [["app-home"]], decls: 4, vars: 0, consts: [[1, "welcome"]], template: function HomeComponent_Template(rf, ctx) { if (rf & 1) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "p", 0);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1, "Welcome to Syndi!");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](2, "p");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](3, "This sample demonstrates how to configure MSAL Angular to login, logout, protect a route, and acquire an access token for a protected resource such as the Microsoft Graph.");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        } }, styles: ["\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IiIsImZpbGUiOiJzcmMvYXBwL2NvbXBvbmVudHMvaG9tZS9ob21lLmNvbXBvbmVudC5jc3MifQ== */"] });
    return HomeComponent;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](HomeComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-home',
                templateUrl: './home.component.html',
                styleUrls: ['./home.component.css']
            }]
    }], function () { return []; }, null); })();


/***/ }),

/***/ "DZ0t":
/*!*********************************************************!*\
  !*** ./src/app/components/profile/profile.component.ts ***!
  \*********************************************************/
/*! exports provided: ProfileComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ProfileComponent", function() { return ProfileComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var _services_graph_service__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../../services/graph.service */ "Tv8v");



var ProfileComponent = /** @class */ (function () {
    function ProfileComponent(graph) {
        this.graph = graph;
    }
    ProfileComponent.prototype.ngOnInit = function () {
        var _this = this;
        this.graph.getProfile()
            .then(function (user) { return _this.profile = user; })
            .catch(function (reason) { return console.error(reason); });
    };
    ProfileComponent.ɵfac = function ProfileComponent_Factory(t) { return new (t || ProfileComponent)(_angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdirectiveInject"](_services_graph_service__WEBPACK_IMPORTED_MODULE_1__["GraphService"])); };
    ProfileComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: ProfileComponent, selectors: [["app-profile"]], decls: 2, vars: 2, template: function ProfileComponent_Template(rf, ctx) { if (rf & 1) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "p");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        } if (rf & 2) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtextInterpolate2"](" Welcome ", ctx.profile == null ? null : ctx.profile.displayName, "(", ctx.profile == null ? null : ctx.profile.mail, ")!\n");
        } }, styles: ["\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IiIsImZpbGUiOiJzcmMvYXBwL2NvbXBvbmVudHMvcHJvZmlsZS9wcm9maWxlLmNvbXBvbmVudC5jc3MifQ== */"] });
    return ProfileComponent;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](ProfileComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-profile',
                templateUrl: './profile.component.html',
                styleUrls: ['./profile.component.css']
            }]
    }], function () { return [{ type: _services_graph_service__WEBPACK_IMPORTED_MODULE_1__["GraphService"] }]; }, null); })();


/***/ }),

/***/ "Sy1n":
/*!**********************************!*\
  !*** ./src/app/app.component.ts ***!
  \**********************************/
/*! exports provided: AppComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppComponent", function() { return AppComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var _services_authentication_service__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./services/authentication.service */ "ej43");
/* harmony import */ var _services_graph_service__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./services/graph.service */ "Tv8v");
/* harmony import */ var _angular_material_toolbar__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @angular/material/toolbar */ "/t3+");
/* harmony import */ var _angular_common__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @angular/common */ "ofXK");
/* harmony import */ var _angular_material_form_field__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @angular/material/form-field */ "kmnG");
/* harmony import */ var _angular_material_input__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @angular/material/input */ "qFsG");
/* harmony import */ var _angular_material_button__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! @angular/material/button */ "bTqV");
/* harmony import */ var _angular_router__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! @angular/router */ "tyNb");










function AppComponent_mat_form_field_4_Template(rf, ctx) { if (rf & 1) {
    var _r6 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵgetCurrentView"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "mat-form-field");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](1, "input", 6);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](2, "button", 7);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function AppComponent_mat_form_field_4_Template_button_click_2_listener() { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵrestoreView"](_r6); var ctx_r5 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵnextContext"](); return ctx_r5.invite(); });
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](3, "Invite");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} if (rf & 2) {
    var ctx_r0 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵnextContext"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("value", ctx_r0.mail);
} }
function AppComponent_mat_form_field_5_button_1_Template(rf, ctx) { if (rf & 1) {
    var _r9 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵgetCurrentView"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "button", 7);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function AppComponent_mat_form_field_5_button_1_Template_button_click_0_listener() { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵrestoreView"](_r9); var ctx_r8 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵnextContext"](2); return ctx_r8.calendar(); });
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1, "Calendar");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} }
var _c0 = function () { return ["profile"]; };
function AppComponent_mat_form_field_5_Template(rf, ctx) { if (rf & 1) {
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "mat-form-field");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](1, AppComponent_mat_form_field_5_button_1_Template, 2, 0, "button", 4);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](2, "a", 8);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](3, "Profile");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} if (rf & 2) {
    var ctx_r1 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵnextContext"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngIf", ctx_r1.loggedIn);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("routerLink", _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵpureFunction0"](2, _c0));
} }
function AppComponent_button_6_Template(rf, ctx) { if (rf & 1) {
    var _r11 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵgetCurrentView"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "button", 7);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function AppComponent_button_6_Template_button_click_0_listener() { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵrestoreView"](_r11); var ctx_r10 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵnextContext"](); return ctx_r10.login(); });
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1, "Login");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} }
function AppComponent_button_7_Template(rf, ctx) { if (rf & 1) {
    var _r13 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵgetCurrentView"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "button", 7);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function AppComponent_button_7_Template_button_click_0_listener() { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵrestoreView"](_r13); var ctx_r12 = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵnextContext"](); return ctx_r12.logout(); });
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1, "Logout");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} }
function AppComponent_router_outlet_9_Template(rf, ctx) { if (rf & 1) {
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](0, "router-outlet");
} }
var AppComponent = /** @class */ (function () {
    function AppComponent(auth, graph) {
        var _this = this;
        this.auth = auth;
        this.graph = graph;
        this.title = 'Syndi';
        // mail = 'hakam@spectory.com';
        this.mail = 'eladt.pro@gmail.com';
        this.isIframe = false;
        this.loggedIn = false;
        this.login = function () { return _this.auth.login(); };
    }
    ;
    AppComponent.prototype.ngOnInit = function () {
        var _this = this;
        this.isIframe = window !== window.parent && !window.opener;
        this.checkAccount();
        this.auth.logged().subscribe(function (account) { return _this.checkAccount(); });
    };
    AppComponent.prototype.checkAccount = function () {
        this.loggedIn = this.auth.authenticated();
    };
    AppComponent.prototype.logout = function () {
        this.auth.logout();
    };
    AppComponent.prototype.invite = function () {
        this.graph.sendInvitation(this.mail);
    };
    AppComponent.prototype.calendar = function () {
        this.graph.getCalendarView('2020-11-01T19:00:00-08:00', '2020-11-02T19:00:00-08:00', '');
    };
    AppComponent.ɵfac = function AppComponent_Factory(t) { return new (t || AppComponent)(_angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdirectiveInject"](_services_authentication_service__WEBPACK_IMPORTED_MODULE_1__["AuthenticationService"]), _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdirectiveInject"](_services_graph_service__WEBPACK_IMPORTED_MODULE_2__["GraphService"])); };
    AppComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: AppComponent, selectors: [["app-root"]], decls: 10, vars: 6, consts: [["color", "primary"], ["href", "/", 1, "title"], [1, "toolbar-spacer"], [4, "ngIf"], ["mat-raised-button", "", 3, "click", 4, "ngIf"], [1, "container"], ["matInput", "", "ng-model", "mail", "placeholder", "guest@mail.com", 3, "value"], ["mat-raised-button", "", 3, "click"], ["mat-button", "", 3, "routerLink"]], template: function AppComponent_Template(rf, ctx) { if (rf & 1) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "mat-toolbar", 0);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](1, "a", 1);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](2);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](3, "div", 2);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](4, AppComponent_mat_form_field_4_Template, 4, 1, "mat-form-field", 3);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](5, AppComponent_mat_form_field_5_Template, 4, 3, "mat-form-field", 3);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](6, AppComponent_button_6_Template, 2, 0, "button", 4);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](7, AppComponent_button_7_Template, 2, 0, "button", 4);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](8, "div", 5);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](9, AppComponent_router_outlet_9_Template, 1, 0, "router-outlet", 3);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        } if (rf & 2) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](2);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtextInterpolate"](ctx.title);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](2);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngIf", ctx.loggedIn);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngIf", ctx.loggedIn);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngIf", !ctx.loggedIn);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngIf", ctx.loggedIn);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](2);
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngIf", !ctx.isIframe);
        } }, directives: [_angular_material_toolbar__WEBPACK_IMPORTED_MODULE_3__["MatToolbar"], _angular_common__WEBPACK_IMPORTED_MODULE_4__["NgIf"], _angular_material_form_field__WEBPACK_IMPORTED_MODULE_5__["MatFormField"], _angular_material_input__WEBPACK_IMPORTED_MODULE_6__["MatInput"], _angular_material_button__WEBPACK_IMPORTED_MODULE_7__["MatButton"], _angular_material_button__WEBPACK_IMPORTED_MODULE_7__["MatAnchor"], _angular_router__WEBPACK_IMPORTED_MODULE_8__["RouterLinkWithHref"], _angular_router__WEBPACK_IMPORTED_MODULE_8__["RouterOutlet"]], styles: [".toolbar-spacer[_ngcontent-%COMP%] {\r\n  -ms-flex: 1 1 auto;\r\n      flex: 1 1 auto;\r\n}\r\n\r\na.title[_ngcontent-%COMP%] {\r\n  color: white;\r\n}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbInNyYy9hcHAvYXBwLmNvbXBvbmVudC5jc3MiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IkFBQUE7RUFDRSxrQkFBYztNQUFkLGNBQWM7QUFDaEI7O0FBRUE7RUFDRSxZQUFZO0FBQ2QiLCJmaWxlIjoic3JjL2FwcC9hcHAuY29tcG9uZW50LmNzcyIsInNvdXJjZXNDb250ZW50IjpbIi50b29sYmFyLXNwYWNlciB7XHJcbiAgZmxleDogMSAxIGF1dG87XHJcbn1cclxuXHJcbmEudGl0bGUge1xyXG4gIGNvbG9yOiB3aGl0ZTtcclxufVxyXG4iXX0= */"] });
    return AppComponent;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](AppComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-root',
                templateUrl: './app.component.html',
                styleUrls: ['./app.component.css']
            }]
    }], function () { return [{ type: _services_authentication_service__WEBPACK_IMPORTED_MODULE_1__["AuthenticationService"] }, { type: _services_graph_service__WEBPACK_IMPORTED_MODULE_2__["GraphService"] }]; }, null); })();


/***/ }),

/***/ "Tv8v":
/*!*******************************************!*\
  !*** ./src/app/services/graph.service.ts ***!
  \*******************************************/
/*! exports provided: GraphService */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "GraphService", function() { return GraphService; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "mrSG");
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var _microsoft_microsoft_graph_client__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/microsoft-graph-client */ "XZ4c");
/* harmony import */ var _models_user__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../models/user */ "2hxB");
/* harmony import */ var _authentication_service__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./authentication.service */ "ej43");






var GraphService = /** @class */ (function () {
    function GraphService(auth) {
        var _this = this;
        this.auth = auth;
        this.client = _microsoft_microsoft_graph_client__WEBPACK_IMPORTED_MODULE_2__["Client"].init({
            debugLogging: true,
            authProvider: function (done) { return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(_this, void 0, void 0, function () {
                return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                    this.auth.getAccessToken()
                        .then(function (token) { return done(null, token); })
                        .catch(function (reason) { return done(reason, null); });
                    return [2 /*return*/];
                });
            }); }
        });
    }
    GraphService.prototype.getProfile = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var graphUser, user;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.auth.authenticated())
                            return [2 /*return*/, null];
                        return [4 /*yield*/, this.client
                                .api('/me')
                                // .select('displayName,mail,mailboxSettings,userPrincipalName')
                                .select('displayName,mail')
                                .get()];
                    case 1:
                        graphUser = _a.sent();
                        user = new _models_user__WEBPACK_IMPORTED_MODULE_3__["User"](graphUser);
                        return [2 /*return*/, user];
                }
            });
        });
    };
    // https://docs.microsoft.com/en-us/graph/api/invitation-post?view=graph-rest-1.0&tabs=http
    GraphService.prototype.sendInvitation = function (mail) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.sendInvitationGraph(mail)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    GraphService.prototype.sendInvitationGraph = function (mail) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var request, invitation;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        request = {
                            "invitedUserEmailAddress": mail,
                            "inviteRedirectUrl": "http://localhost:4200/invitation",
                            // "invitedUserDisplayName": "Elad Tal",
                            // "invitedUserMessageInfo": { "@odata.type": "microsoft.graph.invitedUserMessageInfo" },
                            "sendInvitationMessage": true,
                            "inviteRedeemUrl": "http://localhost:4200/auth",
                        };
                        return [4 /*yield*/, this.client.api('/invitations').post(request)
                                .then(function (result) { return console.log(result); })
                                .catch(function (error) { return console.error(error); })];
                    case 1:
                        invitation = _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    // POST https://graph.microsoft.com/v1.0/invitations
    // Content-type: application/json
    // Content-length: 551
    // {
    //   "invitedUserEmailAddress": "yyy@test.com",
    //   "inviteRedirectUrl": "https://myapp.contoso.com"
    // }
    GraphService.prototype.getCalendarView = function (start, end, timeZone) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var result, error_1;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.auth.authenticated())
                            return [2 /*return*/, null];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.client
                                .api('/me/calendarview')
                                // .header('Prefer', `outlook.timezone="${timeZone}"`)
                                .query({
                                startDateTime: start,
                                endDateTime: end
                            })
                                .select('subject,organizer,start,end')
                                .orderby('start/dateTime')
                                .top(50)
                                .get()];
                    case 2:
                        result = _a.sent();
                        return [2 /*return*/, result.value];
                    case 3:
                        error_1 = _a.sent();
                        console.error('Could not get events', JSON.stringify(error_1, null, 2));
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // <AddEventSnippet>
    GraphService.prototype.addEventToCalendar = function (newEvent) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var error_2;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        // POST /me/events
                        return [4 /*yield*/, this.client
                                .api('/me/events')
                                .post(newEvent)];
                    case 1:
                        // POST /me/events
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_2 = _a.sent();
                        throw Error(JSON.stringify(error_2, null, 2));
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GraphService.ɵfac = function GraphService_Factory(t) { return new (t || GraphService)(_angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵɵinject"](_authentication_service__WEBPACK_IMPORTED_MODULE_4__["AuthenticationService"])); };
    GraphService.ɵprov = _angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵɵdefineInjectable"]({ token: GraphService, factory: GraphService.ɵfac, providedIn: 'root' });
    return GraphService;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵsetClassMetadata"](GraphService, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_1__["Injectable"],
        args: [{
                providedIn: 'root'
            }]
    }], function () { return [{ type: _authentication_service__WEBPACK_IMPORTED_MODULE_4__["AuthenticationService"] }]; }, null); })();


/***/ }),

/***/ "ZAI4":
/*!*******************************!*\
  !*** ./src/app/app.module.ts ***!
  \*******************************/
/*! exports provided: AppModule */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppModule", function() { return AppModule; });
/* harmony import */ var _angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/platform-browser */ "jhN1");
/* harmony import */ var _angular_platform_browser_animations__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @angular/platform-browser/animations */ "R1ws");
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var _angular_material_input__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @angular/material/input */ "qFsG");
/* harmony import */ var _angular_material_form_field__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @angular/material/form-field */ "kmnG");
/* harmony import */ var _angular_material_button__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @angular/material/button */ "bTqV");
/* harmony import */ var _angular_material_list__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @angular/material/list */ "MutI");
/* harmony import */ var _angular_material_toolbar__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! @angular/material/toolbar */ "/t3+");
/* harmony import */ var _app_routing_module__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./app-routing.module */ "vY5A");
/* harmony import */ var _app_component__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./app.component */ "Sy1n");
/* harmony import */ var msal__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! msal */ "YjPl");
/* harmony import */ var _azure_msal_angular__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! @azure/msal-angular */ "E8bv");
/* harmony import */ var _angular_common_http__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! @angular/common/http */ "tk/3");
/* harmony import */ var _components_profile_profile_component__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ./components/profile/profile.component */ "DZ0t");
/* harmony import */ var _components_home_home_component__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ./components/home/home.component */ "BuFo");
/* harmony import */ var _components_logged_out_logged_out_component__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ./components/logged-out/logged-out.component */ "q4k/");
/* harmony import */ var _components_invitation_invitation_component__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! ./components/invitation/invitation.component */ "xr8A");



















var isIE = window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1;
var AppModule = /** @class */ (function () {
    function AppModule() {
    }
    AppModule.ɵmod = _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵɵdefineNgModule"]({ type: AppModule, bootstrap: [_app_component__WEBPACK_IMPORTED_MODULE_9__["AppComponent"]] });
    AppModule.ɵinj = _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵɵdefineInjector"]({ factory: function AppModule_Factory(t) { return new (t || AppModule)(); }, providers: [
            {
                provide: _angular_common_http__WEBPACK_IMPORTED_MODULE_12__["HTTP_INTERCEPTORS"],
                useClass: _azure_msal_angular__WEBPACK_IMPORTED_MODULE_11__["MsalInterceptor"],
                multi: true
            }
        ], imports: [[
                _angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__["BrowserModule"],
                _app_routing_module__WEBPACK_IMPORTED_MODULE_8__["AppRoutingModule"],
                _angular_platform_browser_animations__WEBPACK_IMPORTED_MODULE_1__["BrowserAnimationsModule"],
                _angular_common_http__WEBPACK_IMPORTED_MODULE_12__["HttpClientModule"],
                _angular_material_toolbar__WEBPACK_IMPORTED_MODULE_7__["MatToolbarModule"],
                _angular_material_input__WEBPACK_IMPORTED_MODULE_3__["MatInputModule"],
                _angular_material_form_field__WEBPACK_IMPORTED_MODULE_4__["MatFormFieldModule"],
                _angular_material_button__WEBPACK_IMPORTED_MODULE_5__["MatButtonModule"],
                _angular_material_list__WEBPACK_IMPORTED_MODULE_6__["MatListModule"],
                _app_routing_module__WEBPACK_IMPORTED_MODULE_8__["AppRoutingModule"],
                _azure_msal_angular__WEBPACK_IMPORTED_MODULE_11__["MsalModule"].forRoot({
                    auth: {
                        clientId: '838d477f-42c9-4701-9490-eb6fef98bee4',
                        authority: 'https://login.microsoftonline.com/03a0f57d-d561-42fe-a38b-cd46ed17d941/',
                        redirectUri: 'http://localhost:4200/auth',
                        postLogoutRedirectUri: "http://localhost:4200/goodbye",
                        navigateToLoginRequestUrl: true
                    },
                    cache: {
                        cacheLocation: 'sessionStorage',
                        storeAuthStateInCookie: isIE,
                    },
                    system: {
                        logger: new msal__WEBPACK_IMPORTED_MODULE_10__["Logger"](function (level, message, containsPii) {
                            if (containsPii) {
                                return;
                            }
                            switch (level) {
                                case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Error:
                                    console.error(message);
                                    return;
                                case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Info:
                                    console.info(message);
                                    return;
                                case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Verbose:
                                    console.debug(message);
                                    return;
                                case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Warning:
                                    console.warn(message);
                                    return;
                            }
                        }, {
                            level: msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Info,
                            piiLoggingEnabled: false,
                            correlationId: "[MsalAngular(" + msal__WEBPACK_IMPORTED_MODULE_10__["CryptoUtils"].createNewGuid() + ")]"
                        })
                    }
                }, {
                    popUp: !isIE,
                    consentScopes: [
                        // 'user.read',
                        'offline_access',
                        // 'email',
                        'openid',
                        'profile',
                        'User.Invite.All',
                        'User.ReadWrite.All',
                        'Directory.ReadWrite.All',
                    ],
                    unprotectedResources: [],
                    protectedResourceMap: [
                    // [environment.graph, ['user.read']]
                    ],
                    extraQueryParameters: {}
                })
            ]] });
    return AppModule;
}());

(function () { (typeof ngJitMode === "undefined" || ngJitMode) && _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵɵsetNgModuleScope"](AppModule, { declarations: [_app_component__WEBPACK_IMPORTED_MODULE_9__["AppComponent"],
        _components_profile_profile_component__WEBPACK_IMPORTED_MODULE_13__["ProfileComponent"],
        _components_home_home_component__WEBPACK_IMPORTED_MODULE_14__["HomeComponent"],
        _components_logged_out_logged_out_component__WEBPACK_IMPORTED_MODULE_15__["LoggedOutComponent"],
        _components_invitation_invitation_component__WEBPACK_IMPORTED_MODULE_16__["InvitationComponent"]], imports: [_angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__["BrowserModule"],
        _app_routing_module__WEBPACK_IMPORTED_MODULE_8__["AppRoutingModule"],
        _angular_platform_browser_animations__WEBPACK_IMPORTED_MODULE_1__["BrowserAnimationsModule"],
        _angular_common_http__WEBPACK_IMPORTED_MODULE_12__["HttpClientModule"],
        _angular_material_toolbar__WEBPACK_IMPORTED_MODULE_7__["MatToolbarModule"],
        _angular_material_input__WEBPACK_IMPORTED_MODULE_3__["MatInputModule"],
        _angular_material_form_field__WEBPACK_IMPORTED_MODULE_4__["MatFormFieldModule"],
        _angular_material_button__WEBPACK_IMPORTED_MODULE_5__["MatButtonModule"],
        _angular_material_list__WEBPACK_IMPORTED_MODULE_6__["MatListModule"],
        _app_routing_module__WEBPACK_IMPORTED_MODULE_8__["AppRoutingModule"], _azure_msal_angular__WEBPACK_IMPORTED_MODULE_11__["MsalModule"]] }); })();
/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵsetClassMetadata"](AppModule, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_2__["NgModule"],
        args: [{
                declarations: [
                    _app_component__WEBPACK_IMPORTED_MODULE_9__["AppComponent"],
                    _components_profile_profile_component__WEBPACK_IMPORTED_MODULE_13__["ProfileComponent"],
                    _components_home_home_component__WEBPACK_IMPORTED_MODULE_14__["HomeComponent"],
                    _components_logged_out_logged_out_component__WEBPACK_IMPORTED_MODULE_15__["LoggedOutComponent"],
                    _components_invitation_invitation_component__WEBPACK_IMPORTED_MODULE_16__["InvitationComponent"]
                ],
                imports: [
                    _angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__["BrowserModule"],
                    _app_routing_module__WEBPACK_IMPORTED_MODULE_8__["AppRoutingModule"],
                    _angular_platform_browser_animations__WEBPACK_IMPORTED_MODULE_1__["BrowserAnimationsModule"],
                    _angular_common_http__WEBPACK_IMPORTED_MODULE_12__["HttpClientModule"],
                    _angular_material_toolbar__WEBPACK_IMPORTED_MODULE_7__["MatToolbarModule"],
                    _angular_material_input__WEBPACK_IMPORTED_MODULE_3__["MatInputModule"],
                    _angular_material_form_field__WEBPACK_IMPORTED_MODULE_4__["MatFormFieldModule"],
                    _angular_material_button__WEBPACK_IMPORTED_MODULE_5__["MatButtonModule"],
                    _angular_material_list__WEBPACK_IMPORTED_MODULE_6__["MatListModule"],
                    _app_routing_module__WEBPACK_IMPORTED_MODULE_8__["AppRoutingModule"],
                    _azure_msal_angular__WEBPACK_IMPORTED_MODULE_11__["MsalModule"].forRoot({
                        auth: {
                            clientId: '838d477f-42c9-4701-9490-eb6fef98bee4',
                            authority: 'https://login.microsoftonline.com/03a0f57d-d561-42fe-a38b-cd46ed17d941/',
                            redirectUri: 'http://localhost:4200/auth',
                            postLogoutRedirectUri: "http://localhost:4200/goodbye",
                            navigateToLoginRequestUrl: true
                        },
                        cache: {
                            cacheLocation: 'sessionStorage',
                            storeAuthStateInCookie: isIE,
                        },
                        system: {
                            logger: new msal__WEBPACK_IMPORTED_MODULE_10__["Logger"](function (level, message, containsPii) {
                                if (containsPii) {
                                    return;
                                }
                                switch (level) {
                                    case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Error:
                                        console.error(message);
                                        return;
                                    case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Info:
                                        console.info(message);
                                        return;
                                    case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Verbose:
                                        console.debug(message);
                                        return;
                                    case msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Warning:
                                        console.warn(message);
                                        return;
                                }
                            }, {
                                level: msal__WEBPACK_IMPORTED_MODULE_10__["LogLevel"].Info,
                                piiLoggingEnabled: false,
                                correlationId: "[MsalAngular(" + msal__WEBPACK_IMPORTED_MODULE_10__["CryptoUtils"].createNewGuid() + ")]"
                            })
                        }
                    }, {
                        popUp: !isIE,
                        consentScopes: [
                            // 'user.read',
                            'offline_access',
                            // 'email',
                            'openid',
                            'profile',
                            'User.Invite.All',
                            'User.ReadWrite.All',
                            'Directory.ReadWrite.All',
                        ],
                        unprotectedResources: [],
                        protectedResourceMap: [
                        // [environment.graph, ['user.read']]
                        ],
                        extraQueryParameters: {}
                    })
                ],
                providers: [
                    {
                        provide: _angular_common_http__WEBPACK_IMPORTED_MODULE_12__["HTTP_INTERCEPTORS"],
                        useClass: _azure_msal_angular__WEBPACK_IMPORTED_MODULE_11__["MsalInterceptor"],
                        multi: true
                    }
                ],
                bootstrap: [_app_component__WEBPACK_IMPORTED_MODULE_9__["AppComponent"]]
            }]
    }], null, null); })();


/***/ }),

/***/ "ej43":
/*!****************************************************!*\
  !*** ./src/app/services/authentication.service.ts ***!
  \****************************************************/
/*! exports provided: AuthenticationService */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AuthenticationService", function() { return AuthenticationService; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var rxjs__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! rxjs */ "qCKp");
/* harmony import */ var _azure_msal_angular__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @azure/msal-angular */ "E8bv");




var AuthParams = {
    redirectUri: 'http://localhost:4200/auth',
    scopes: [
        // "user.read",
        // "mailboxsettings.read",
        // "calendars.readwrite",
        'User.Invite.All',
        'User.ReadWrite.All',
        'Directory.ReadWrite.All',
    ]
};
var AuthenticationService = /** @class */ (function () {
    function AuthenticationService(msal, broadcastService) {
        var _this = this;
        this.msal = msal;
        this.broadcastService = broadcastService;
        this.logged$ = new rxjs__WEBPACK_IMPORTED_MODULE_1__["BehaviorSubject"](null);
        this.isIE = function () { return window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1; };
        this.getAccount = function () { return _this.msal.getAccount(); };
        this.authenticated = function () { return !!_this.getAccount(); };
        this.login = function () {
            if (_this.isIE())
                _this.msal.loginRedirect();
            else
                _this.msal.loginPopup()
                    .catch(function (reason) {
                    console.error('Login failed', JSON.stringify(reason, null, 2));
                });
        };
        this.logout = function () { return _this.msal.logout(); };
        this.getAccessToken = function () {
            return new Promise(function (resolve, reject) {
                _this.msal.acquireTokenSilent(AuthParams)
                    .then(function (result) { return resolve(result.accessToken); })
                    .catch(function (error) {
                    //Acquire token silent failure, and send an interactive request
                    console.error(error.errorMessage, error);
                    if (error.errorMessage.indexOf("interaction_required") === -1)
                        reject(error);
                    if (_this.isIE())
                        _this.msal.acquireTokenRedirect(AuthParams);
                    else
                        _this.msal.acquireTokenPopup(AuthParams)
                            .then(function (response) { return resolve(response.accessToken); })
                            .catch(reject);
                });
            });
        };
        this.msal.handleRedirectCallback(function (authError, response) {
            if (authError) {
                console.error('Redirect Error: ', authError.errorMessage);
                return;
            }
            console.log('Redirect Success: ', response.accessToken);
        });
        this.broadcastService.subscribe('msal:loginSuccess', function () {
            _this.logged$.next(_this.getAccount());
        });
    }
    AuthenticationService.prototype.logged = function () {
        return this.logged$.asObservable();
    };
    AuthenticationService.ɵfac = function AuthenticationService_Factory(t) { return new (t || AuthenticationService)(_angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵinject"](_azure_msal_angular__WEBPACK_IMPORTED_MODULE_2__["MsalService"]), _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵinject"](_azure_msal_angular__WEBPACK_IMPORTED_MODULE_2__["BroadcastService"])); };
    AuthenticationService.ɵprov = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineInjectable"]({ token: AuthenticationService, factory: AuthenticationService.ɵfac, providedIn: 'root' });
    return AuthenticationService;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](AuthenticationService, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Injectable"],
        args: [{
                providedIn: 'root'
            }]
    }], function () { return [{ type: _azure_msal_angular__WEBPACK_IMPORTED_MODULE_2__["MsalService"] }, { type: _azure_msal_angular__WEBPACK_IMPORTED_MODULE_2__["BroadcastService"] }]; }, null); })();


/***/ }),

/***/ "q4k/":
/*!***************************************************************!*\
  !*** ./src/app/components/logged-out/logged-out.component.ts ***!
  \***************************************************************/
/*! exports provided: LoggedOutComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "LoggedOutComponent", function() { return LoggedOutComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");


var LoggedOutComponent = /** @class */ (function () {
    function LoggedOutComponent() {
    }
    LoggedOutComponent.prototype.ngOnInit = function () {
    };
    LoggedOutComponent.ɵfac = function LoggedOutComponent_Factory(t) { return new (t || LoggedOutComponent)(); };
    LoggedOutComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: LoggedOutComponent, selectors: [["app-logged-out"]], decls: 2, vars: 0, template: function LoggedOutComponent_Template(rf, ctx) { if (rf & 1) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "p");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1, "logged-out works!");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        } }, styles: ["\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IiIsImZpbGUiOiJzcmMvYXBwL2NvbXBvbmVudHMvbG9nZ2VkLW91dC9sb2dnZWQtb3V0LmNvbXBvbmVudC5jc3MifQ== */"] });
    return LoggedOutComponent;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](LoggedOutComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-logged-out',
                templateUrl: './logged-out.component.html',
                styleUrls: ['./logged-out.component.css']
            }]
    }], function () { return []; }, null); })();


/***/ }),

/***/ "vY5A":
/*!***************************************!*\
  !*** ./src/app/app-routing.module.ts ***!
  \***************************************/
/*! exports provided: AppRoutingModule */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppRoutingModule", function() { return AppRoutingModule; });
/* harmony import */ var _components_invitation_invitation_component__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./components/invitation/invitation.component */ "xr8A");
/* harmony import */ var _components_logged_out_logged_out_component__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./components/logged-out/logged-out.component */ "q4k/");
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var _angular_router__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @angular/router */ "tyNb");
/* harmony import */ var _components_profile_profile_component__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./components/profile/profile.component */ "DZ0t");
/* harmony import */ var _azure_msal_angular__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @azure/msal-angular */ "E8bv");
/* harmony import */ var _components_home_home_component__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./components/home/home.component */ "BuFo");









var routes = [
    { path: 'profile', component: _components_profile_profile_component__WEBPACK_IMPORTED_MODULE_4__["ProfileComponent"], canActivate: [_azure_msal_angular__WEBPACK_IMPORTED_MODULE_5__["MsalGuard"]] },
    { path: 'auth', redirectTo: 'profile' },
    { path: 'invitation', component: _components_invitation_invitation_component__WEBPACK_IMPORTED_MODULE_0__["InvitationComponent"] },
    { path: 'goodbye', component: _components_logged_out_logged_out_component__WEBPACK_IMPORTED_MODULE_1__["LoggedOutComponent"] },
    { path: '', component: _components_home_home_component__WEBPACK_IMPORTED_MODULE_6__["HomeComponent"] }
];
var AppRoutingModule = /** @class */ (function () {
    function AppRoutingModule() {
    }
    AppRoutingModule.ɵmod = _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵɵdefineNgModule"]({ type: AppRoutingModule });
    AppRoutingModule.ɵinj = _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵɵdefineInjector"]({ factory: function AppRoutingModule_Factory(t) { return new (t || AppRoutingModule)(); }, imports: [[_angular_router__WEBPACK_IMPORTED_MODULE_3__["RouterModule"].forRoot(routes, { useHash: false })], _angular_router__WEBPACK_IMPORTED_MODULE_3__["RouterModule"]] });
    return AppRoutingModule;
}());

(function () { (typeof ngJitMode === "undefined" || ngJitMode) && _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵɵsetNgModuleScope"](AppRoutingModule, { imports: [_angular_router__WEBPACK_IMPORTED_MODULE_3__["RouterModule"]], exports: [_angular_router__WEBPACK_IMPORTED_MODULE_3__["RouterModule"]] }); })();
/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_2__["ɵsetClassMetadata"](AppRoutingModule, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_2__["NgModule"],
        args: [{
                imports: [_angular_router__WEBPACK_IMPORTED_MODULE_3__["RouterModule"].forRoot(routes, { useHash: false })],
                exports: [_angular_router__WEBPACK_IMPORTED_MODULE_3__["RouterModule"]]
            }]
    }], null, null); })();


/***/ }),

/***/ "xr8A":
/*!***************************************************************!*\
  !*** ./src/app/components/invitation/invitation.component.ts ***!
  \***************************************************************/
/*! exports provided: InvitationComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "InvitationComponent", function() { return InvitationComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");


var InvitationComponent = /** @class */ (function () {
    function InvitationComponent() {
    }
    InvitationComponent.prototype.ngOnInit = function () {
    };
    InvitationComponent.ɵfac = function InvitationComponent_Factory(t) { return new (t || InvitationComponent)(); };
    InvitationComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: InvitationComponent, selectors: [["app-invitation"]], decls: 2, vars: 0, template: function InvitationComponent_Template(rf, ctx) { if (rf & 1) {
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "p");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](1, "invitation works!");
            _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        } }, styles: ["\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IiIsImZpbGUiOiJzcmMvYXBwL2NvbXBvbmVudHMvaW52aXRhdGlvbi9pbnZpdGF0aW9uLmNvbXBvbmVudC5jc3MifQ== */"] });
    return InvitationComponent;
}());

/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](InvitationComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-invitation',
                templateUrl: './invitation.component.html',
                styleUrls: ['./invitation.component.css']
            }]
    }], function () { return []; }, null); })();


/***/ }),

/***/ "zUnb":
/*!*********************!*\
  !*** ./src/main.ts ***!
  \*********************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "fXoL");
/* harmony import */ var _environments_environment__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./environments/environment */ "AytR");
/* harmony import */ var _app_app_module__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./app/app.module */ "ZAI4");
/* harmony import */ var _angular_platform_browser__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @angular/platform-browser */ "jhN1");




if (_environments_environment__WEBPACK_IMPORTED_MODULE_1__["environment"].production) {
    Object(_angular_core__WEBPACK_IMPORTED_MODULE_0__["enableProdMode"])();
}
_angular_platform_browser__WEBPACK_IMPORTED_MODULE_3__["platformBrowser"]().bootstrapModule(_app_app_module__WEBPACK_IMPORTED_MODULE_2__["AppModule"])
    .catch(function (err) { return console.error(err); });


/***/ }),

/***/ "zn8P":
/*!******************************************************!*\
  !*** ./$$_lazy_route_resource lazy namespace object ***!
  \******************************************************/
/*! no static exports found */
/***/ (function(module, exports) {

function webpackEmptyAsyncContext(req) {
	// Here Promise.resolve().then() is used instead of new Promise() to prevent
	// uncaught exception popping up in devtools
	return Promise.resolve().then(function() {
		var e = new Error("Cannot find module '" + req + "'");
		e.code = 'MODULE_NOT_FOUND';
		throw e;
	});
}
webpackEmptyAsyncContext.keys = function() { return []; };
webpackEmptyAsyncContext.resolve = webpackEmptyAsyncContext;
module.exports = webpackEmptyAsyncContext;
webpackEmptyAsyncContext.id = "zn8P";

/***/ })

},[[0,"runtime","vendor"]]]);
//# sourceMappingURL=main.js.map