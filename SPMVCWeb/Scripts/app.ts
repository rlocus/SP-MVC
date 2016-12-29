/// <reference path="typings/angularjs/angular.d.ts" />
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
    private _scriptPromises;
    static SharePointAppName = "SharePointApp";
    static SPServiceName = "SPService";

    constructor() {
        this._initialized = false;
        this._scriptPromises = {};
    }

    public init(preloadedScripts: any[]) {
        var self = this;
        if (preloadedScripts) {
            let $ = preloadedScripts["jquery"];
            self.$ = $;
            if (self.$) {
                (<any>self.$).cachedScript = function (url, options) {
                    options = self.$.extend(options || {}, {
                        dataType: "script",
                        cache: true,
                        url: url
                    });
                    return self.$.ajax(options);
                };
            } else {
                throw "jQuery is not loaded!";
            }

            let $angular = preloadedScripts["angular"];
            self.$angular = $angular;
            if (!self.$angular) {
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
        this.spApp = this.$angular.module(App.SharePointAppName, [
            //'ngSanitize',
            'officeuifabric.core',
            'officeuifabric.components'
        ]).service(App.SPServiceName, function ($http, $q) {
            this.getFormDigest = () => {
                var deferred = self.$.Deferred();
                var url = $pnp.util.combinePaths(self.appWebUrl, "_api/contextinfo")
                var executor = new SP.RequestExecutor(self.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    url: url,
                    method: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        var formDigestValue = JSON.parse(<string>data.body).d.GetContextWebInformation.FormDigestValue;
                        deferred.resolve(formDigestValue);
                    },
                    error: function (error) {
                        deferred.reject(error);
                    }
                });
                return deferred.promise();
            }
        }).filter('unsafe', function ($sce) {
            return $sce.trustAsHtml;
        }).directive('compile', [
            '$compile', ($compile) => {
                return (scope, element, attrs) => {
                    scope.$watch(
                        (scope) => {
                            // watch the 'compile' expression for changes
                            return scope.$eval((<any>attrs).compile);
                        },
                        (value) => {
                            // when the 'compile' expression changes
                            // assign it into the current DOM
                            element.html(value);
                            // compile the new DOM and link it to the current
                            // scope.
                            // NOTE: we only compile .childNodes so that
                            // we don't get into infinite loop compiling ourselves
                            $compile(element.contents())(scope);
                        }
                    );
                };
            }
        ]);
        this._initialized = true;
    }

    public ensureScript(url): JQueryXHR {
        if (url) {
            url = url.toLowerCase().replace("~sphost", this.scriptBase);
            var scriptPromise = this._scriptPromises[url];
            if (!scriptPromise) {
                scriptPromise = (<any>this.$).cachedScript(url);
                this._scriptPromises[url] = scriptPromise;
            }
            return scriptPromise;
        }
        return null;
    }

    public delay = (function () {
        var timer = 0;
        return (callback: () => void, ms: number) => {
            clearTimeout(timer);
            timer = setTimeout(callback, ms);
        };
    })();

    public render(...modules: App.IModule[]) {
        var self = this;
        if (!self._initialized) {
            throw "App is not initialized!";
        }
        self.ensureScript("~sphost/MicrosoftAjax.js").then(function () {
            self.ensureScript("~sphost/SP.Runtime.js").then(function () {
                self.ensureScript("~sphost/SP.RequestExecutor.js").then(function () {
                    self.ensureScript("~sphost/SP.js").then(function () {
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
            });
        });
    }

    public get_ListView(options: App.Module.IListViewOptions): App.Module.ListView {
        var self = this;
        return new App.Module.ListView(self, options);
    }

    public get_ListsView(options: App.Module.IListsViewOptions): App.Module.ListsView {
        var self = this;
        return new App.Module.ListsView(self, options);
    }

    public get_BasePermissions(permMask: string): SP.BasePermissions {
        var permMaskHigh = permMask.length <= 10 ? 0 : parseInt(permMask.substring(2, permMask.length - 8), 16);
        var permMaskLow = permMask.length <= 10 ? parseInt(permMask) : parseInt(permMask.substring(permMask.length - 8, permMask.length), 16);
        var permissions = new SP.BasePermissions();
        permissions.initPropertiesFromJson({ "High": permMaskHigh, "Low": permMaskLow });
        return permissions;
    }
}

export = new App();

declare module App {

    export interface ISPService {
        getFormDigest();
    }

    export interface IModuleOptions {
        controllerName: string;
    }

    export interface IModule {
        render();
    }
}

module App.Module {

    export interface IListViewOptions extends App.IModuleOptions {
        listTitle: string;
        listId?: string;
        listUrl?: string;
        viewId?: string;
        viewXml?: string,
        orderBy?: string;
        sortAsc?: boolean;
        filter?: string;
        limit?: number;
        expands?: string[];
        paged?: boolean,
        rootFolder: string,
        //fields?: string[];
        onload: ($scope: ng.IScope) => void;
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

        private getEntity(listItem) {
            var permMask = listItem["PermMask"];
            var permissions = this._app.get_BasePermissions(permMask);
            var $permissions = {
                edit: permissions.has(SP.PermissionKind.editListItems),
                delete: permissions.has(SP.PermissionKind.deleteListItems)
            };
            var $events = { menuOpened: false, isSelected: false };
            return { $data: listItem, $events: $events, $permissions: $permissions };
        }

        private getListItems() {
            var self = this;
            var deferred = self._app.$.Deferred();
            var url: string = null;
            if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                var query;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                    query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists;
                    query.concat("/GetById(@list)")
                    query.query.add("@list", "'" + self._options.listId + "'");
                } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                    query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                    query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                }
                query.concat("/RenderListDataAsStream");
                query.query.add("View", self._options.viewId);
                if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                    query.query.add("SortField", self._options.orderBy);
                }
                if (!$pnp.util.stringIsNullOrEmpty(<any>self._options.sortAsc)) {
                    query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                }
                url = query.toUrlAndQuery();
                var parameters = <any>{ "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": 0 | 1 | 2 /*| 4 | 8*/ };
                if (!$pnp.util.stringIsNullOrEmpty(self._options.viewXml)) {
                    parameters.ViewXml = self._options.viewXml;
                }
                if (!$pnp.util.stringIsNullOrEmpty(<any>self._options.paged)) {
                    parameters.Paging = self._options.paged;
                }
                if (!$pnp.util.stringIsNullOrEmpty(self._options.rootFolder)) {
                    parameters.FolderServerRelativeUrl = self._options.rootFolder;
                }
                var postBody = JSON.stringify({ "parameters": parameters });
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    url: url,
                    method: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    body: postBody,
                    success: function (data) {
                        var result = JSON.parse(<string>data.body);
                        deferred.resolve(result);
                    },
                    error: function (error) {
                        deferred.reject(error);
                    }
                });
                return deferred.promise();
            }
            if (!$pnp.util.stringIsNullOrEmpty(self._options.viewXml)) {
                var query;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                    query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId);
                } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                    query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                    query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                }
                //query.concat("/getitems");
                //url = query.toUrlAndQuery();
                //var postBody = JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": self._options.viewXml } });
                //var executor = new SP.RequestExecutor(self._app.appWebUrl);
                //executor.executeAsync(<SP.RequestInfo>{
                //    url: url,
                //    method: "POST",
                //    body: postBody,
                //    headers: {
                //        "accept": "application/json;odata=verbose",
                //        "content-Type": "application/json;odata=verbose"
                //    },
                //    success: function (data) {
                //        var listItems = JSON.parse(<string>data.body).d.results;
                //        deferred.resolve(listItems);
                //    },
                //    error: function (error) {
                //        deferred.reject(error);
                //    }
                //});
                query.concat("/renderlistdata(@viewXml)");
                query.query.add("@viewXml", "'" + self._options.viewXml + "'");
                url = query.toUrlAndQuery();
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    url: url,
                    method: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        var result = JSON.parse(JSON.parse(<string>data.body).d.RenderListData);
                        deferred.resolve(result);
                    },
                    error: function (error) {
                        deferred.reject(error);
                    }
                });
                return deferred.promise();
            }
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
            self._app.spApp.factory("ListViewFactory", ($q, $http) => {
                var factory = {} as IListViewFactory;
                factory.listItems = [];
                factory.getListItems = () => {
                    var deferred = $q.defer();
                    self.getListItems().then((data: any) => {
                        factory.listItems.splice(0, factory.listItems.length);
                        var rows = data.Row ? data.Row : (data.ListData ? data.ListData.Row : data);
                        self._app.$.each(rows, (function (i, listItem) {
                            factory.listItems.push(self.getEntity(listItem));
                        }));
                        deferred.resolve(factory.listItems);
                    }, deferred.reject);
                    return deferred.promise;
                }
                return factory;
            });
            var deferred = self._app.$.Deferred();
            self._app.spApp.controller(self._options.controllerName, [
                '$scope', 'ListViewFactory', App.SPServiceName, function ($scope: ng.IScope, factory: IListViewFactory, service: App.ISPService) {
                    (<any>$scope).loading = true;
                    (<any>$scope).listItems = [];
                    factory.getListItems().then(() => {
                        (<any>$scope).listItems = self._app.$angular.copy(factory.listItems);
                        (<any>$scope).loading = false;
                        deferred.resolve();
                    }, deferred.reject);
                    (<any>$scope).selection = {
                        commandBar: {
                            searchTerm: null,
                            createEnabled: false,
                            viewEnabled: false,
                            deleteEnabled: false,
                            view: function (listItem) {
                                if (!listItem) {
                                    var selectedItems = (<any>$scope).table.selectedItems;
                                    listItem = self._app.$(selectedItems).get(0);
                                }
                                if (listItem) {
                                }
                            },
                            delete: function (listItem) {
                                if (!listItem) {
                                    var selectedItems = (<any>$scope).table.selectedItems;

                                }
                            },
                            clearSelection: function () {
                                var selectedItems = (<any>$scope).table.selectedItems;
                                if (selectedItems.length > 0) {
                                    self._app.$.each((<any>$scope).table.rows, (i, item) => {
                                        if (item.selected) {
                                            item.selected = false;
                                        }
                                    });
                                }
                            }
                        },
                        openMenu: (listItem) => {
                            if (listItem) {
                                if (!listItem.$events.menuOpened) {
                                    self._app.$.each((<any>$scope).listItems, ((i, listItem) => {
                                        listItem.$events.menuOpened = false;
                                    }));
                                }
                                listItem.$events.menuOpened = !listItem.$events.menuOpened;
                            }
                        }
                    };
                    $scope.$watch('table.selectedItems', function (newValue: Array<any>, oldValue: Array<any>) {
                        (<any>$scope).selection.commandBar.viewEnabled = newValue.length === 1;
                        (<any>$scope).selection.commandBar.deleteEnabled = newValue.length > 0;
                        (<any>$scope).selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                    }, true);
                    $scope.$watch('selection.commandBar.searchTerm', function (newValue: string, oldValue: string) {
                        //(<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                        //self._app.delay(() => {
                        //    $scope.$apply(function () {

                        //    });
                        //}, self._options.delay);
                    }, false);

                    if (typeof self._options.onload === "function") {
                        self._options.onload($scope);
                    }
                }
            ]);
            return deferred.promise();
        }
    }

    interface IListViewFactory {
        listItems: Array<any>;
        getListItems();
    }

    interface IListsViewFactory {
        lists: Array<any>;
        getLists();
        updateList(list, service: App.ISPService);
    }

    export interface IListsViewOptions extends App.IModuleOptions {
        delay: number;
    }

    export class ListsView implements App.IModule {
        private _options: IListsViewOptions;
        private _app: App;

        constructor(app: App, options: IListsViewOptions) {
            if (!app) {
                throw "App must be specified for ListView!";
            }
            this._app = app;
            this._options = this._app.$.extend(true, { delay: 1000 }, options);
        }

        public getLists() {
            var self = this;
            var deferred = self._app.$.Deferred();
            var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
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

        public getList(listId) {
            var self = this;
            var deferred = self._app.$.Deferred();
            var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
            var executor = new SP.RequestExecutor(self._app.appWebUrl);
            executor.executeAsync(<SP.RequestInfo>{
                url: url,
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-Type": "application/json;odata=verbose"
                },
                success: function (data) {
                    var list = JSON.parse(<string>data.body).d;
                    deferred.resolve(list);
                },
                error: function (error) {
                    deferred.reject(error);
                }
            });
            return deferred.promise();
        }

        public updateList(listId, properties, digestValue) {
            var self = this;
            var deferred = self._app.$.Deferred();
            var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).toUrlAndQuery();
            var body = JSON.stringify($pnp.util.extend({
                "__metadata": { "type": "SP.List" },
            }, properties));

            var executor = new SP.RequestExecutor(self._app.appWebUrl);
            executor.executeAsync(<SP.RequestInfo>{
                body: body,
                url: url,
                method: "POST",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-Type": "application/json;odata=verbose",
                    "IF-Match": "*",
                    "X-HTTP-Method": "MERGE",
                    "X-RequestDigest": digestValue
                },
                success: function (data) {
                    deferred.resolve();
                },
                error: function (error) {
                    deferred.reject(error);
                }
            });
            return deferred.promise();
        }

        private getEntity(list) {
            switch (list.BaseType) {
                case 1:
                    list.Type = "Document Library";
                    break;
                default:
                    list.Type = "List";
                    break;
            }
            var permissions = new SP.BasePermissions();
            permissions.initPropertiesFromJson(list["EffectiveBasePermissions"]);
            var $permissions = {
                manage: permissions.has(SP.PermissionKind.manageLists)
            }
            var $events = { menuOpened: false, isSelected: false };
            return { $data: list, $events: $events, $permissions: $permissions };
        }

        public render() {
            var self = this;
            self._app.spApp.factory("ListsViewFactory", ($q, $http) => {
                var factory = {} as IListsViewFactory;
                factory.lists = [];
                factory.getLists = () => {
                    var deferred = $q.defer();
                    self.getLists().then((data: Array<any>) => {
                        factory.lists.splice(0, factory.lists.length);
                        self._app.$.each(data, (function (i, list) {
                            if (!list.Hidden) {
                                var entity = self.getEntity(list);
                                factory.lists.push(entity);
                            }
                        }));
                        deferred.resolve(factory.lists);
                    }, deferred.reject);
                    return deferred.promise;
                }
                factory.updateList = (list, service: App.ISPService) => {
                    var deferred = $q.defer();
                    service.getFormDigest().done((digestValue: string) => {
                        var properties = {
                            'Title': list.Title,
                            'Description': list.Description
                        };
                        self.updateList(list.Id, properties, digestValue).done(() => {
                            self.getList(list.Id).done((data) => {
                                deferred.resolve(data);
                            });
                        }).fail(deferred.reject);
                    }).fail(deferred.reject);
                    return deferred.promise;
                };
                return factory;
            });
            var deferred = self._app.$.Deferred();
            self._app.spApp.controller(self._options.controllerName, ['$scope', 'ListsViewFactory', App.SPServiceName, function ($scope: ng.IScope, factory: IListsViewFactory, service: App.ISPService) {
                (<any>$scope).loading = true;
                (<any>$scope).lists = [];
                factory.getLists().then(() => {
                    (<any>$scope).lists = self._app.$angular.copy(factory.lists);
                    (<any>$scope).loading = false;
                    deferred.resolve();
                }, () => {
                    (<any>$scope).loading = false;
                    deferred.reject();
                });
                (<any>$scope).selection = {
                    settings: {
                        opened: false,
                        data: [],
                        editMode: false,
                        onEdit: () => {
                            (<any>$scope).selection.settings.editMode = true;
                        },
                        onSave: () => {
                            return factory.updateList((<any>$scope).selection.settings.data, service).then((data) => {
                                (<any>$scope).selection.settings.data = self._app.$.extend(true, {}, data);
                                (<any>$scope).selection.settings.editMode = false;
                                var entity = self.getEntity(data);
                                $.each((<any>$scope).lists, (i, list) => {
                                    if (list.$data.Id === entity.$data.Id) {
                                        list.$data = entity.$data;
                                        list.$permissions = entity.$permissions;
                                    }
                                });
                            });
                        }
                    },
                    commandBar: {
                        searchTerm: null,
                        createEnabled: false,
                        viewEnabled: false,
                        deleteEnabled: false,
                        settingsEnabled: false,
                        openSettings: function (list) {
                            if (!list) {
                                var selectedItems = (<any>$scope).table.selectedItems;
                                list = self._app.$(selectedItems).get(0);
                            }
                            if (list) {
                                if (!(<any>$scope).selection.settings.opened) {
                                    (<any>$scope).selection.settings.editMode = false;
                                    (<any>$scope).selection.settings.canEdit = list.$permissions.manage;
                                    (<any>$scope).selection.settings.data.Id = list.$data.Id;
                                    (<any>$scope).selection.settings.data.Title = list.$data.Title;
                                    (<any>$scope).selection.settings.data.Description = list.$data.Description;
                                } else {
                                    (<any>$scope).selection.settings.data = [];
                                }
                                (<any>$scope).selection.settings.opened = !(<any>$scope).selection.settings.opened;
                            }
                        },
                        view: function (list) {
                            if (!list) {
                                var selectedItems = (<any>$scope).table.selectedItems;
                                list = self._app.$(selectedItems).get(0);
                            }
                            if (list) {
                                window.location.href = "/Home/List?ListId=" + list.$data.Id + "&SPHostUrl=" + decodeURIComponent(self._app.hostWebUrl) + "&SPAppWebUrl=" + decodeURIComponent(self._app.appWebUrl);
                            }
                        },
                        delete: function (list) {
                            if (!list) {
                                var selectedItems = (<any>$scope).table.selectedItems;

                            }
                        },
                        clearSelection: function () {
                            var selectedItems = (<any>$scope).table.selectedItems;
                            if (selectedItems.length > 0) {
                                self._app.$.each((<any>$scope).table.rows, (i, item) => {
                                    if (item.selected) {
                                        item.selected = false;
                                    }
                                });
                            }
                        }
                    },
                    openMenu: (list) => {
                        if (list) {
                            if (!list.$events.menuOpened) {
                                self._app.$.each((<any>$scope).lists, ((i, list) => {
                                    list.$events.menuOpened = false;
                                }));
                            }
                            list.$events.menuOpened = !list.$events.menuOpened;
                        }
                    }
                };
                $scope.$watch('table.selectedItems', function (newValue: Array<any>, oldValue: Array<any>) {
                    (<any>$scope).selection.commandBar.viewEnabled = newValue.length === 1;
                    (<any>$scope).selection.commandBar.deleteEnabled = newValue.length > 0;
                    (<any>$scope).selection.commandBar.settingsEnabled = newValue.length === 1;
                    (<any>$scope).selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                }, true);
                $scope.$watch('selection.commandBar.searchTerm', function (newValue: string, oldValue: string) {
                    (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                    self._app.delay(() => {
                        $scope.$apply(function () {
                            var lists;
                            if (newValue && newValue !== oldValue) {
                                lists = []
                                self._app.$.each(factory.lists, (i, list) => {
                                    if ((<any>list).$data && new RegExp(newValue, 'i').test((<any>list).$data.Title)) {
                                        lists.push(list);
                                    }
                                });
                            } else {
                                lists = factory.lists;
                            }
                            (<any>$scope).lists = self._app.$angular.copy(lists);
                        });
                    }, self._options.delay);
                }, false);
            }]);
            return deferred.promise();
        }
    }
}