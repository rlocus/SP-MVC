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
            var self = this;
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
            ]).service(App.SPServiceName, function ($http, $q) {
                this.getFormDigest = function () {
                    var deferred = self.$.Deferred();
                    var url = $pnp.util.combinePaths(self.appWebUrl, "_api/contextinfo");
                    var executor = new SP.RequestExecutor(self.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            var formDigestValue = JSON.parse(data.body).d.GetContextWebInformation.FormDigestValue;
                            deferred.resolve(formDigestValue);
                        },
                        error: function (error) {
                            deferred.reject(error);
                        }
                    });
                    return deferred.promise();
                };
            });
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
        App.SPServiceName = "SPService";
        return App;
    }());
    var App;
    (function (App) {
        var Module;
        (function (Module) {
            var delay = (function () {
                var timer = 0;
                return function (callback, ms) {
                    clearTimeout(timer);
                    timer = setTimeout(callback, ms);
                };
            })();
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
                ListsView.prototype.getList = function (listId) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            var list = JSON.parse(data.body).d;
                            deferred.resolve(list);
                        },
                        error: function (error) {
                            deferred.reject(error);
                        }
                    });
                    return deferred.promise();
                };
                ListsView.prototype.updateList = function (listId, properties, digestValue) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).toUrlAndQuery();
                    var body = JSON.stringify($pnp.util.extend({
                        "__metadata": { "type": "SP.List" },
                    }, properties));
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
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
                };
                ListsView.prototype.getEntity = function (list) {
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
                    var $events = { menuOpened: false, isSelected: false };
                    return { $data: list, $events: $events, $permissions: $permissions };
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
                                self._app.$.each(data, (function (i, list) {
                                    if (!list.Hidden) {
                                        var entity = self.getEntity(list);
                                        factory.lists.push(entity);
                                    }
                                }));
                                deferred.resolve(factory.lists);
                            }, deferred.reject);
                            return deferred.promise;
                        };
                        factory.updateList = function (list, service) {
                            var deferred = $q.defer();
                            service.getFormDigest().done(function (digestValue) {
                                var properties = {
                                    'Title': list.Title,
                                    'Description': list.Description
                                };
                                self.updateList(list.Id, properties, digestValue).done(function () {
                                    self.getList(list.Id).done(function (data) {
                                        deferred.resolve(data);
                                    });
                                }).fail(deferred.reject);
                            }).fail(deferred.reject);
                            return deferred.promise;
                        };
                        return factory;
                    });
                    var deferred = self._app.$.Deferred();
                    self._app.spApp.controller(self._options.controllerName, ['$scope', 'ListsViewFactory', App.SPServiceName, function ($scope, factory, service) {
                            $scope.loading = true;
                            $scope.lists = factory.lists;
                            factory.getLists().then(function () {
                                $scope.loading = false;
                                deferred.resolve();
                            }, function () {
                                $scope.loading = false;
                                deferred.reject();
                            });
                            $scope.settingsOpened = false;
                            $scope.selection = {
                                settings: {
                                    data: [],
                                    editMode: false,
                                    onEdit: function () {
                                        $scope.selection.settings.editMode = true;
                                    },
                                    onSave: function () {
                                        return factory.updateList($scope.selection.settings.data, service).then(function (data) {
                                            $scope.selection.settings.data = self._app.$.extend(true, {}, data);
                                            $scope.selection.settings.editMode = false;
                                            var entity = self.getEntity(data);
                                            $.each($scope.lists, (function (i, list) {
                                                if (list.$data.Id === entity.$data.Id) {
                                                    list.$data = entity.$data;
                                                    list.$permissions = entity.$permissions;
                                                }
                                            }));
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
                                            var selectedItems = $scope.table.selectedItems;
                                            list = self._app.$(selectedItems).get(0);
                                        }
                                        if (list) {
                                            if (!$scope.settingsOpened) {
                                                $scope.selection.settings.editMode = false;
                                                $scope.selection.settings.canEdit = list.$permissions.manage;
                                                $scope.selection.settings.data.Id = list.$data.Id;
                                                $scope.selection.settings.data.Title = list.$data.Title;
                                                $scope.selection.settings.data.Description = list.$data.Description;
                                            }
                                            else {
                                                $scope.selection.settings.data = [];
                                            }
                                            $scope.settingsOpened = !$scope.settingsOpened;
                                        }
                                    },
                                    view: function (list) {
                                        if (!list) {
                                            var selectedItems = $scope.table.selectedItems;
                                            list = self._app.$(selectedItems).get(0);
                                        }
                                    },
                                    delete: function (list) {
                                        if (!list) {
                                            var selectedItems = $scope.table.selectedItems;
                                        }
                                    },
                                    clearSelection: function () {
                                        var selectedItems = $scope.table.selectedItems;
                                        if (selectedItems.length > 0) {
                                            self._app.$.each($scope.table.rows, (function (i, item) {
                                                if (item.selected) {
                                                    item.selected = false;
                                                }
                                            }));
                                        }
                                    }
                                }
                            };
                            $scope.$watch('table.selectedItems', function (newValue, oldValue) {
                                $scope.selection.commandBar.viewEnabled = newValue.length === 1;
                                $scope.selection.commandBar.deleteEnabled = newValue.length > 0;
                                $scope.selection.commandBar.settingsEnabled = newValue.length === 1;
                                $scope.selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                            }, true);
                            $scope.$watch('selection.commandBar.searchTerm', function (newValue, oldValue) {
                                if (newValue && newValue !== oldValue) {
                                    delay(function () {
                                        $scope.$apply(function () {
                                            $scope.lists = self._app.$.grep(factory.lists, function (list) {
                                                return list.$data && new RegExp(newValue, 'i').test(list.$data.Title) /*(<any>list).$data.Title.toUpperCase().startsWith(newValue.toUpperCase())*/;
                                            });
                                        });
                                    }, 1000);
                                }
                                else {
                                    $scope.lists = factory.lists;
                                }
                            }, true);
                            $scope.openMenu = function (list) {
                                if (list) {
                                    if (!list.$events.menuOpened) {
                                        self._app.$.each($scope.lists, (function (i, list) {
                                            list.$events.menuOpened = false;
                                        }));
                                    }
                                    list.$events.menuOpened = !list.$events.menuOpened;
                                }
                            };
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
