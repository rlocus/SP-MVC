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

    static SharePointAppName = "SharePointApp";
    static SPServiceName = "SPService";

    constructor() {
        this._initialized = false;
    }

    public init(preloadedScripts: any[]) {
        var self = this;
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

        this.spApp = this.$angular.module(App.SharePointAppName, [
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
        });
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
            self.ensureScript(self.scriptBase + "/sp.runtime.js").then(function (data) {
                self.ensureScript(self.scriptBase + "/SP.RequestExecutor.js").then(function (data) {
                    self.ensureScript(self.scriptBase + "/SP.js").then(function (data) {
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

    public get_Lists(options: App.IModuleOptions): App.Module.ListsView {
        var self = this;
        return new App.Module.ListsView(self, options);
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
        //constructor(app: App, options: App.IModuleOptions);
        render();
    }
}

module App.Module {
    var delay = (function () {
        var timer = 0;
        return (callback: () => void, ms: number) => {
            clearTimeout(timer);
            timer = setTimeout(callback, ms);
        };
    })();

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

    interface IListsViewFactory {
        lists: any;
        getLists();
        updateList(list, service: App.ISPService);
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
                    self._app.$.each(factory.lists, (i, list) => {
                        (<any>$scope).lists.push(list);
                    });
                    (<any>$scope).loading = false;
                    deferred.resolve();
                }, () => {
                    (<any>$scope).loading = false;
                    deferred.reject();
                });
                (<any>$scope).settingsOpened = false;
                (<any>$scope).selection = {
                    settings: {
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
                                if (!(<any>$scope).settingsOpened) {
                                    (<any>$scope).selection.settings.editMode = false;
                                    (<any>$scope).selection.settings.canEdit = list.$permissions.manage;
                                    (<any>$scope).selection.settings.data.Id = list.$data.Id;
                                    (<any>$scope).selection.settings.data.Title = list.$data.Title;
                                    (<any>$scope).selection.settings.data.Description = list.$data.Description;
                                } else {
                                    (<any>$scope).selection.settings.data = [];
                                }
                                (<any>$scope).settingsOpened = !(<any>$scope).settingsOpened;
                            }
                        },
                        view: function (list) {
                            if (!list) {
                                var selectedItems = (<any>$scope).table.selectedItems;
                                list = self._app.$(selectedItems).get(0);
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
                                //self._app.$.each((<any>$scope).rows, (i, item) => {
                                //    if (item.selected) {
                                //        item.selected = false;
                                //    }
                                //});
                                self._app.$.each((<any>$scope).table.rows, (i, item) => {
                                    if (item.selected) {
                                        item.selected = false;
                                    }
                                });
                            }
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
                    //(<any>$scope).selection.commandBar.clearSelection();
                    //(<any>$scope).lists = [];
                    (<any>$scope).lists.splice(0, (<any>$scope).lists.length);
                    (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                    //(<any>$scope).rows = (<any>$scope).table.rows = [];
                    if (newValue && newValue !== oldValue) {
                        delay(() => {
                            $scope.$apply(function() {
                                self._app.$.each(factory.lists, (i, list) => {
                                    if ((<any>list).$data && new RegExp(newValue, 'i').test((<any>list).$data.Title)) {
                                        (<any>$scope).lists.push(list);
                                    }
                                });
                            });
                        }, 1000);
                    } else {
                        self._app.$.each(factory.lists, (i, list) => {
                            (<any>$scope).lists.push(list);
                        });
                    }
                }, false);
                (<any>$scope).openMenu = function (list) {
                    if (list) {
                        if (!list.$events.menuOpened) {
                            self._app.$.each((<any>$scope).lists, (function (i, list) {
                                list.$events.menuOpened = false;
                            }));
                        }
                        list.$events.menuOpened = !list.$events.menuOpened;
                    }
                };
            }]);
            return deferred.promise();
        }
    }
}