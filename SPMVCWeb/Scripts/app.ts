﻿/// <reference path="typings/angularjs/angular.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />

import * as $pnp from "pnp";
import * as $ from "jquery";
//import * as $angular from "angular";

class App {

    public appWebUrl: string;
    public hostWebUrl: string;
    public scriptBase: string;
    public $: JQueryStatic;
    public $angular: angular.IAngularStatic;
    public spApp: ng.IModule;

    private _initialized: boolean;

    static SharePointAppName = "SharePointApp";

    constructor() {
        this._initialized = false;
    }

    public init(preloadedScripts: any[]) {
        if (preloadedScripts) {
            let $ = preloadedScripts["jquery"];
            this.$ = $;
            if (this.$) {
                (<any>this.$).cachedScript = function (url, options) {
                    options = $.extend(options || {}, {
                        dataType: "script",
                        cache: true,
                        url: url
                    });
                    return jQuery.ajax(options);
                };
            } else {
                throw "jQuery is not loaded!";
            }

            let $angular = preloadedScripts["angular"];
            this.$angular = $angular;
            if (!this.$angular) {
                throw "Angular is not loaded!";
            }
        }
        this.hostWebUrl = $pnp.util.getUrlParamByName("SPHostUrl");
        if ($pnp.util.stringIsNullOrEmpty(this.hostWebUrl)) {
            throw "SPHostUrl url parameter must be specified!";
        }
        this.appWebUrl = $pnp.util.getUrlParamByName("SPAppWebUrl");
        if ($pnp.util.stringIsNullOrEmpty(this.appWebUrl)) {
            throw "SPAppWebUrl url parameter must be specified!";
        }
        this.scriptBase = $pnp.util.combinePaths(this.hostWebUrl, "_layouts/15");
        this.spApp = this.$angular.module(App.SharePointAppName, []);
        this._initialized = true;
    }

    public ensureScript(url): JQueryXHR {
        return (<any>this.$).cachedScript(url);
    }

    public render(...modules: App.IModule[]) {
        var self = this;
        if (!self._initialized) {
            throw "App is not initialized!";
        }

        self.ensureScript(self.scriptBase + "/MicrosoftAjax.js").then(function (data) {
            self.ensureScript(self.scriptBase + "/SP.RequestExecutor.js").then(function (data) {

                if ($pnp.util.isArray(modules)) {
                    self.$.each(modules, (i: number, module: App.IModule) => {
                        module.render();
                    });
                }

                self.$angular.element(function () {
                    self.$angular.bootstrap(document, [App.SharePointAppName]);
                });
            });
        });
    }

    public get_ListView(options: App.Module.IListViewOptions): App.Module.ListView {
        var self = this;
        return new App.Module.ListView(self, options);
    }

    public get_Lists(options: App.IModuleOptions): App.Module.ListsView {
        var self = this;
        return new App.Module.ListsView(self, options);
    }
}

export = new App();

declare module App {

    export interface IModuleOptions {
        controllerName: string;
    }

    export interface IModule {
        //constructor(app: App, options: App.IModuleOptions);
        render();
    }
}

module App.Module {

    export interface IListViewOptions extends App.IModuleOptions {
        listTitle: string;
        listId?: string;
        listUrl?: string;
        orderBy?: string;
        sortAsc?: boolean;
        filter?: string;
        limit?: number;
        expands?: string[];
    }

    export class ListView implements App.IModule {
        private _options: IListViewOptions;
        private _app: App;

        constructor(app: App, options: App.Module.IListViewOptions) {
            if (!app) {
                throw "App must be specified for ListView!";
            }
            this._app = app;
            this._options = options;
        }

        private getItems() {
            var self = this;
            var deferred = self._app.$.Deferred();
            var url: string = null;
            if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                var items = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId).items;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                    items = items.orderBy(self._options.orderBy, self._options.sortAsc);
                }
                if (!$pnp.util.stringIsNullOrEmpty(self._options.filter)) {
                    items = items.filter(self._options.filter);
                }
                if (self._options.limit > 0) {
                    items = items.top(self._options.limit);
                }
                if ($pnp.util.isArray(self._options.expands)) {
                    items = items.expand(<any>self._options.expands);
                }
                url = items.toUrlAndQuery();
            } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                var items = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl).items;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                    items = items.orderBy(self._options.orderBy, self._options.sortAsc);
                }
                if (!$pnp.util.stringIsNullOrEmpty(self._options.filter)) {
                    items = items.filter(self._options.filter);
                }
                if (self._options.limit > 0) {
                    items = items.top(self._options.limit);
                }
                if ($pnp.util.isArray(self._options.expands)) {
                    items = items.expand(<any>self._options.expands);
                }
                url = items.toUrlAndQuery();
            } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                var items = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listTitle).items;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                    items = items.orderBy(self._options.orderBy, self._options.sortAsc);
                }
                if (!$pnp.util.stringIsNullOrEmpty(self._options.filter)) {
                    items = items.filter(self._options.filter);
                }
                if (self._options.limit > 0) {
                    items = items.top(self._options.limit);
                }
                if ($pnp.util.isArray(self._options.expands)) {
                    items = items.expand(<any>self._options.expands);
                }
                url = items.toUrlAndQuery();
            }
            if (url !== null) {
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        var listItems = JSON.parse(<string>data.body).d.results;
                        deferred.resolve(listItems);
                    },
                    error: function (error) {
                        deferred.reject(error);
                    }
                });
            }
            return deferred.promise();
        }

        public render() {
            var self = this;
            var deferred = self._app.$.Deferred();
            self._app.spApp.controller(self._options.controllerName, function ($scope: ng.IScope) {
                self.getItems().then((listItems) => {
                    (<any>$scope).listItems = listItems;
                    $scope.$apply();
                    deferred.resolve();
                }, deferred.reject);
            });
            return deferred.promise();
        }
    }

    export class ListsView implements App.IModule {
        private _options: App.IModuleOptions;
        private _app: App;

        constructor(app: App, options: App.IModuleOptions) {
            if (!app) {
                throw "App must be specified for ListView!";
            }
            this._app = app;
            this._options = options;
        }

        public getLists() {
            var self = this;
            var deferred = self._app.$.Deferred();
            var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.toUrlAndQuery();
            var executor = new SP.RequestExecutor(self._app.appWebUrl);
            executor.executeAsync(<SP.RequestInfo>{
                url: url,
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-Type": "application/json;odata=verbose"
                },
                success: function (data) {
                    var lists = JSON.parse(<string>data.body).d.results;
                    deferred.resolve(lists);
                },
                error: function (error) {
                    deferred.reject(error);
                }
            });
            return deferred.promise();
        }

        public render() {
            var self = this;
            var deferred = self._app.$.Deferred();
            self._app.spApp.controller(self._options.controllerName, function ($scope: ng.IScope) {
                self.getLists().then((lists) => {
                    (<any>$scope).lists = lists;
                    $scope.$apply();
                    deferred.resolve();
                }, deferred.reject);
            });
            return deferred.promise();
        }
    }
}