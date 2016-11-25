/// <reference path="typings/angularjs/angular.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
define(["require", "exports", "pnp", "jquery"], function (require, exports, $pnp, $) {
    "use strict";
    //import * as $angular from "angular";
    var App = (function () {
        function App() {
            this._initialized = false;
        }
        App.prototype.init = function (preloadedScripts) {
            if (preloadedScripts) {
                var $_1 = preloadedScripts["jquery"];
                this.$ = $_1;
                if (this.$) {
                    this.$.cachedScript = function (url, options) {
                        options = $_1.extend(options || {}, {
                            dataType: "script",
                            cache: true,
                            url: url
                        });
                        return jQuery.ajax(options);
                    };
                }
                else {
                    throw "jQuery is not loaded!";
                }
                var $angular = preloadedScripts["angular"];
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
            ]);
            this._initialized = true;
        };
        App.prototype.ensureScript = function (url) {
            return this.$.cachedScript(url);
        };
        App.prototype.render = function () {
            var modules = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                modules[_i - 0] = arguments[_i];
            }
            var self = this;
            if (!self._initialized) {
                throw "App is not initialized!";
            }
            self.ensureScript(self.scriptBase + "/MicrosoftAjax.js").then(function (data) {
                self.ensureScript(self.scriptBase + "/sp.runtime.js").then(function (data) {
                    self.ensureScript(self.scriptBase + "/SP.RequestExecutor.js").then(function (data) {
                        self.ensureScript(self.scriptBase + "/SP.js").then(function (data) {
                            if ($pnp.util.isArray(modules)) {
                                self.$.each(modules, function (i, module) {
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
        };
        App.prototype.get_ListView = function (options) {
            var self = this;
            return new App.Module.ListView(self, options);
        };
        App.prototype.get_Lists = function (options) {
            var self = this;
            return new App.Module.ListsView(self, options);
        };
        App.SharePointAppName = "SharePointApp";
        return App;
    }());
    var App;
    (function (App) {
        var Module;
        (function (Module) {
            var ListView = (function () {
                function ListView(app, options) {
                    if (!app) {
                        throw "App must be specified for ListView!";
                    }
                    this._app = app;
                    this._options = options;
                }
                ListView.prototype.getItems = function () {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    var url = null;
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
                            items = items.expand(self._options.expands);
                        }
                        url = items.toUrlAndQuery();
                    }
                    else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
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
                            items = items.expand(self._options.expands);
                        }
                        url = items.toUrlAndQuery();
                    }
                    else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
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
                            items = items.expand(self._options.expands);
                        }
                        url = items.toUrlAndQuery();
                    }
                    if (url !== null) {
                        var executor = new SP.RequestExecutor(self._app.appWebUrl);
                        executor.executeAsync({
                            url: url,
                            method: "GET",
                            headers: {
                                "accept": "application/json;odata=verbose",
                                "content-Type": "application/json;odata=verbose"
                            },
                            success: function (data) {
                                var listItems = JSON.parse(data.body).d.results;
                                deferred.resolve(listItems);
                            },
                            error: function (error) {
                                deferred.reject(error);
                            }
                        });
                    }
                    return deferred.promise();
                };
                ListView.prototype.render = function () {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    self._app.spApp.controller(self._options.controllerName, function ($scope) {
                        self.getItems().then(function (listItems) {
                            $scope.listItems = listItems;
                            $scope.$apply();
                            deferred.resolve();
                        }, deferred.reject);
                    });
                    return deferred.promise();
                };
                return ListView;
            }());
            Module.ListView = ListView;
            var ListsView = (function () {
                function ListsView(app, options) {
                    if (!app) {
                        throw "App must be specified for ListView!";
                    }
                    this._app = app;
                    this._options = options;
                }
                ListsView.prototype.getLists = function () {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            var lists = JSON.parse(data.body).d.results;
                            deferred.resolve(lists);
                        },
                        error: function (error) {
                            deferred.reject(error);
                        }
                    });
                    return deferred.promise();
                };
                ListsView.prototype.render = function () {
                    var self = this;
                    self._app.spApp.factory("ListsViewFactory", function ($q, $http) {
                        var factory = {};
                        factory.lists = [];
                        factory.getLists = function () {
                            var deferred = $q.defer();
                            self.getLists().then(function (data) {
                                factory.lists.splice(0, factory.lists.length);
                                $.each(data, (function (i, list) {
                                    if (!list.Hidden) {
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
                                        };
                                        var $events = { menuOpened: false, delete: $permissions.manage ? '' : 'disabled' };
                                        factory.lists.push({ $data: list, $events: $events, $permissions: $permissions });
                                    }
                                }));
                                deferred.resolve(data);
                            }, deferred.reject);
                            return deferred.promise;
                        };
                        return factory;
                    });
                    var deferred = self._app.$.Deferred();
                    self._app.spApp.controller(self._options.controllerName, ['$scope', 'ListsViewFactory', function ($scope, factory) {
                            $scope.lists = factory.lists;
                            $scope.settingsOpened = false;
                            $scope.selected = {
                                settings: null
                            };
                            $scope.openMenu = function (list) {
                                if (!list.$events.menuOpened) {
                                    $.each($scope.lists, (function (i, list) {
                                        list.$events.menuOpened = false;
                                    }));
                                }
                                list.$events.menuOpened = !list.$events.menuOpened;
                            };
                            $scope.openSettings = function (list) {
                                if (!$scope.settingsOpened) {
                                    $scope.selected.settings = list.$data;
                                }
                                else {
                                    $scope.selected.settings = null;
                                }
                                $scope.settingsOpened = !$scope.settingsOpened;
                            };
                            $scope.viewList = function (list) {
                            };
                            factory.getLists().then(deferred.resolve, deferred.reject);
                        }]);
                    return deferred.promise();
                };
                return ListsView;
            }());
            Module.ListsView = ListsView;
        })(Module = App.Module || (App.Module = {}));
    })(App || (App = {}));
    return new App();
});
