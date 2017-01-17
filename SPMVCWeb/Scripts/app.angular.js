/// <reference path="typings/angularjs/angular.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
/// <reference path="typings/microsoft-ajax/microsoft.ajax.d.ts" />
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
define(["require", "exports", "pnp", "jquery", "./app.module"], function (require, exports, $pnp, $, app) {
    "use strict";
    var Angular;
    (function (Angular) {
        "use strict";
        var App = (function (_super) {
            __extends(App, _super);
            function App() {
                _super.apply(this, arguments);
            }
            App.prototype.init = function (preloadedScripts) {
                var self = this;
                var $angular = preloadedScripts["angular"];
                self.$angular = $angular;
                if (!self.$angular) {
                    throw "angular is not loaded!";
                }
                self.ngModule = self.$angular.module(App.SharePointAppName, [
                    //'ngSanitize',
                    "officeuifabric.core",
                    "officeuifabric.components"
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
                            error: function (data, errorCode, errorMessage) {
                                if (data.body) {
                                    try {
                                        var error = JSON.parse(data.body);
                                        if (error && error.error) {
                                            errorMessage = error.error.message.value;
                                        }
                                    }
                                    catch (e) {
                                    }
                                }
                                self.$(self).trigger("app-error", [errorMessage]);
                                deferred.reject(data, errorCode, errorMessage);
                            }
                        });
                        return deferred.promise();
                    };
                }).filter("unsafe", function ($sce) {
                    return $sce.trustAsHtml;
                }).directive("compile", [
                    "$compile", function ($compile) {
                        return function (scope, element, attrs) {
                            scope.$watch(function (scope) {
                                // watch the 'compile' expression for changes
                                return scope.$eval(attrs.compile);
                            }, function (value) {
                                // when the 'compile' expression changes
                                // assign it into the current DOM
                                element.html(value);
                                // compile the new DOM and link it to the current
                                // scope.
                                // NOTE: we only compile .childNodes so that
                                // we don't get into infinite loop compiling ourselves
                                $compile(element.contents())(scope);
                            });
                        };
                    }
                ]);
                _super.prototype.init.call(this, preloadedScripts);
            };
            App.prototype.render = function (modules) {
                var self = this;
                self.$(self).on("app-render", function () {
                    self.$angular.element(function () {
                        self.$angular.bootstrap(document, [App.SharePointAppName]);
                    });
                });
                _super.prototype.render.call(this, modules);
            };
            App.prototype.get_ListView = function (options) {
                var self = this;
                //if (!self.is_initialized()) {
                //    throw "App is not initialized!";
                //}
                return new App.Module.ListView(self, options);
            };
            App.prototype.get_ListsView = function (options) {
                var self = this;
                //if (!self.is_initialized()) {
                //    throw "App is not initialized!";
                //}
                return new App.Module.ListsView(self, options);
            };
            App.SharePointAppName = "SharePointApp";
            App.SPServiceName = "SPService";
            return App;
        }(app.App.AppBase));
        Angular.App = App;
        var App;
        (function (App) {
            var Module;
            (function (Module) {
                var ListView = (function (_super) {
                    __extends(ListView, _super);
                    function ListView(app, options) {
                        _super.call(this, app, options);
                    }
                    ListView.prototype.getEntity = function (listItem) {
                        var permMask = listItem["PermMask"];
                        var permissions = this.get_app().get_BasePermissions(permMask);
                        var $permissions = {
                            edit: permissions.has(SP.PermissionKind.editListItems),
                            delete: permissions.has(SP.PermissionKind.deleteListItems)
                        };
                        var $events = { menuOpened: false, isSelected: false };
                        return { $data: listItem, $events: $events, $permissions: $permissions };
                    };
                    ListView.prototype.render = function () {
                        var self = this;
                        var allTokens = [];
                        var app = this.get_app();
                        var options = self.get_options();
                        app.ngModule.factory("ListViewFactory", function ($q, $http) {
                            var factory = {};
                            factory.listItems = [];
                            factory.getToken = function (offset) {
                                if (offset === void 0) { offset = 0; }
                                var token = null;
                                if (offset < 0) {
                                    token = factory.$prevToken;
                                    var skipNext = 1 - offset;
                                    if (allTokens.length > skipNext) {
                                        token = allTokens[(allTokens.length - 1) - skipNext];
                                    }
                                }
                                else if (offset > 0) {
                                    token = factory.$nextToken;
                                    if (offset > 0) {
                                        var index = allTokens.indexOf(factory.$nextToken);
                                        if ((index + offset - 1) < allTokens.length) {
                                            token = allTokens[index + offset - 1];
                                        }
                                    }
                                }
                                else {
                                    token = factory.$currentToken;
                                }
                                return token;
                            };
                            factory.getListItems = function (token) {
                                var deferred = $q.defer();
                                self.getListItems(token).then(function (data) {
                                    //factory.listItems.splice(0, factory.listItems.length);
                                    factory.listItems = [];
                                    if (data.ListData) {
                                        app.$.each(data.ListData.Row, (function (i, listItem) {
                                            factory.listItems.push(self.getEntity(listItem));
                                        }));
                                        factory.$nextToken = data.ListData.NextHref;
                                        factory.$prevToken = data.ListData.PrevHref;
                                        factory.$currentToken = token;
                                        factory.$first = data.ListData.FirstRow ? Number(data.ListData.FirstRow) : 0;
                                        factory.$last = data.ListData.LastRow ? Number(data.ListData.LastRow) : 0;
                                    }
                                    deferred.resolve();
                                }, deferred.reject);
                                return deferred.promise;
                            };
                            return factory;
                        });
                        var deferred = app.$.Deferred();
                        app.ngModule.controller(options.controllerName, [
                            "$scope", "ListViewFactory", App.SPServiceName, function ($scope, factory, service) {
                                $scope.loading = true;
                                $scope.listItems = [];
                                factory.getListItems().then(function () {
                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                    $scope.selection.pager.first = factory.$first;
                                    $scope.selection.pager.last = factory.$last;
                                    $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                    allTokens.push(factory.$nextToken);
                                    $scope.loading = false;
                                    deferred.resolve();
                                }, deferred.reject);
                                $scope.selection = {
                                    commandBar: {
                                        searchTerm: null,
                                        selectionText: "",
                                        createEnabled: false,
                                        viewEnabled: false,
                                        deleteEnabled: false,
                                        view: function (listItem) {
                                            if (!listItem) {
                                                var selectedItems = $scope.table.selectedItems;
                                                listItem = app.$(selectedItems).get(0);
                                            }
                                            if (listItem) {
                                            }
                                        },
                                        delete: function (listItem) {
                                            if (!listItem) {
                                                var selectedItems = $scope.table.selectedItems;
                                            }
                                        },
                                        clearSelection: function () {
                                            var selectedItems = $scope.table.selectedItems;
                                            if (selectedItems.length > 0) {
                                                app.$.each($scope.table.rows, function (i, item) {
                                                    if (item.selected) {
                                                        item.selected = false;
                                                    }
                                                });
                                            }
                                        }
                                    },
                                    openMenu: function (listItem) {
                                        if (listItem) {
                                            if (!listItem.$events.menuOpened) {
                                                app.$.each($scope.listItems, (function (i, listItem) {
                                                    listItem.$events.menuOpened = false;
                                                }));
                                            }
                                            listItem.$events.menuOpened = !listItem.$events.menuOpened;
                                        }
                                    },
                                    pager: {
                                        first: 0,
                                        last: 0,
                                        prevEnabled: false,
                                        nextEnabled: false,
                                        refresh: function () {
                                            if ($scope.loading) {
                                                return;
                                            }
                                            var token = options.appendRows === true ? null : factory.getToken(0);
                                            $scope.table.rows.splice(0, $scope.table.rows.length);
                                            $scope.selection.pager.prevEnabled = false;
                                            $scope.selection.pager.nextEnabled = false;
                                            factory.getListItems(token).then(function () {
                                                //(<any>$scope).selection.commandBar.clearSelection();
                                                if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                }
                                                else {
                                                    $scope.listItems = app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                                }
                                                $scope.selection.pager.first = options.appendRows === true ? ($scope.listItems.length > 0 ? 1 : 0) : factory.$first;
                                                $scope.selection.pager.last = options.appendRows === true ? ($scope.listItems.length) : factory.$last;
                                                $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                $scope.selection.pager.prevEnabled = options.appendRows !== true && !$pnp.util.stringIsNullOrEmpty(factory.getToken(-1));
                                                if (!token) {
                                                    allTokens = [];
                                                }
                                                else {
                                                    allTokens.pop();
                                                }
                                                allTokens.push(factory.$nextToken);
                                                deferred.resolve();
                                            }, deferred.reject);
                                        },
                                        next: function (offset) {
                                            if (offset === void 0) { offset = 1; }
                                            if ($scope.loading) {
                                                return;
                                            }
                                            if (!$scope.selection.pager.nextEnabled)
                                                return;
                                            $scope.table.rows.splice(0, $scope.table.rows.length);
                                            var token = factory.getToken(offset);
                                            $scope.selection.pager.prevEnabled = false;
                                            $scope.selection.pager.nextEnabled = false;
                                            factory.getListItems(token).then(function () {
                                                //(<any>$scope).selection.commandBar.clearSelection();
                                                if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                }
                                                else {
                                                    $scope.listItems = app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                                }
                                                $scope.selection.pager.first = options.appendRows === true ? ($scope.listItems.length > 0 ? 1 : 0) : factory.$first;
                                                $scope.selection.pager.last = options.appendRows === true ? ($scope.listItems.length) : factory.$last;
                                                $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                $scope.selection.pager.prevEnabled = options.appendRows !== true && !$pnp.util.stringIsNullOrEmpty(factory.getToken(-1));
                                                if ($pnp.util.stringIsNullOrEmpty(token)) {
                                                    allTokens = [];
                                                }
                                                allTokens.push(factory.$nextToken);
                                                deferred.resolve();
                                            }, deferred.reject);
                                        },
                                        prev: function (offset) {
                                            if (offset === void 0) { offset = 1; }
                                            if ($scope.loading)
                                                return;
                                            if (!$scope.selection.pager.prevEnabled)
                                                return;
                                            $scope.table.rows.splice(0, $scope.table.rows.length);
                                            offset = Math.min(-1, -offset);
                                            var token = options.appendRows === true ? null : factory.getToken(offset);
                                            $scope.selection.pager.prevEnabled = false;
                                            $scope.selection.pager.nextEnabled = false;
                                            factory.getListItems(token).then(function () {
                                                //(<any>$scope).selection.commandBar.clearSelection();
                                                if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                }
                                                else {
                                                    $scope.listItems = app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                                }
                                                $scope.selection.pager.first = options.appendRows === true ? ($scope.listItems.length > 0 ? 1 : 0) : factory.$first;
                                                $scope.selection.pager.last = options.appendRows === true ? ($scope.listItems.length) : factory.$last;
                                                $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                $scope.selection.pager.prevEnabled = options.appendRows !== true && !$pnp.util.stringIsNullOrEmpty(factory.getToken(-1));
                                                if ($pnp.util.stringIsNullOrEmpty(token)) {
                                                    allTokens = [];
                                                }
                                                else {
                                                    var skipNext = 1 - offset;
                                                    while (skipNext > 0) {
                                                        allTokens.pop();
                                                        skipNext--;
                                                    }
                                                    allTokens.push(factory.$nextToken);
                                                }
                                                deferred.resolve();
                                            }, deferred.reject);
                                        },
                                    }
                                };
                                $scope.$watch("table.selectedItems", function (newValue, oldValue) {
                                    $scope.selection.commandBar.viewEnabled = newValue.length === 1;
                                    $scope.selection.commandBar.deleteEnabled = newValue.length > 0;
                                    $scope.selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                                }, true);
                                $scope.$watch("selection.commandBar.searchTerm", function (newValue, oldValue) {
                                    //(<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                                    //self._app.delay(() => {
                                    //    $scope.$apply(function () {
                                    //    });
                                    //}, self._options.delay);
                                }, false);
                                app.$(self).trigger("model-render", [$scope, factory]);
                            }
                        ]);
                        return deferred.promise();
                    };
                    return ListView;
                }(app.App.Module.ListViewBase));
                Module.ListView = ListView;
                var ListsView = (function (_super) {
                    __extends(ListsView, _super);
                    function ListsView() {
                        _super.apply(this, arguments);
                    }
                    ListsView.prototype.getEntity = function (list) {
                        switch (list.BaseType) {
                            case 0:
                                list.Type = "List";
                                break;
                            case 1:
                                list.Type = "Document Library";
                                break;
                            case 2:
                                list.Type = "Unused";
                                break;
                            case 3:
                                list.Type = "Discussion Board";
                                break;
                            case 4:
                                list.Type = "Survey";
                                break;
                            case 5:
                                list.Type = "Issue";
                                break;
                            default:
                                list.Type = "None";
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
                        var app = this.get_app();
                        var options = self.get_options();
                        app.ngModule.factory("ListsViewFactory", function ($q, $http) {
                            var factory = {};
                            factory.lists = [];
                            factory.getLists = function () {
                                var deferred = $q.defer();
                                self.getLists().then(function (data) {
                                    factory.lists.splice(0, factory.lists.length);
                                    app.$.each(data, (function (i, list) {
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
                                        "Title": list.Title,
                                        "Description": list.Description
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
                        var deferred = app.$.Deferred();
                        app.ngModule.controller(options.controllerName, ["$scope", "ListsViewFactory", App.SPServiceName, function ($scope, factory, service) {
                                $scope.loading = true;
                                $scope.lists = [];
                                factory.getLists().then(function () {
                                    $scope.lists = app.$angular.copy(factory.lists);
                                    $scope.loading = false;
                                    deferred.resolve();
                                }, function () {
                                    $scope.loading = false;
                                    deferred.reject();
                                });
                                $scope.selection = {
                                    settings: {
                                        opened: false,
                                        data: [],
                                        editMode: false,
                                        onEdit: function () {
                                            $scope.selection.settings.editMode = true;
                                        },
                                        onSave: function () {
                                            return factory.updateList($scope.selection.settings.data, service).then(function (data) {
                                                $scope.selection.settings.data = app.$.extend(true, {}, data);
                                                $scope.selection.settings.editMode = false;
                                                var entity = self.getEntity(data);
                                                $.each($scope.lists, function (i, list) {
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
                                        selectionText: "",
                                        createEnabled: false,
                                        viewEnabled: false,
                                        deleteEnabled: false,
                                        settingsEnabled: false,
                                        openSettings: function (list) {
                                            if (!list) {
                                                var selectedItems = $scope.table.selectedItems;
                                                list = app.$(selectedItems).get(0);
                                            }
                                            if (list) {
                                                if (!$scope.selection.settings.opened) {
                                                    $scope.selection.settings.editMode = false;
                                                    $scope.selection.settings.canEdit = list.$permissions.manage;
                                                    $scope.selection.settings.data.Id = list.$data.Id;
                                                    $scope.selection.settings.data.Title = list.$data.Title;
                                                    $scope.selection.settings.data.Description = list.$data.Description;
                                                }
                                                else {
                                                    $scope.selection.settings.data = [];
                                                }
                                                $scope.selection.settings.opened = !$scope.selection.settings.opened;
                                            }
                                        },
                                        view: function (list) {
                                            if (!list) {
                                                var selectedItems = $scope.table.selectedItems;
                                                list = app.$(selectedItems).get(0);
                                            }
                                            if (list) {
                                                window.location.href = "/Home/List?ListId=" + list.$data.Id + "&SPHostUrl=" + decodeURIComponent(app.hostWebUrl) + "&SPAppWebUrl=" + decodeURIComponent(app.appWebUrl);
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
                                                app.$.each($scope.table.rows, function (i, item) {
                                                    if (item.selected) {
                                                        item.selected = false;
                                                    }
                                                });
                                            }
                                        }
                                    },
                                    openMenu: function (list) {
                                        if (list) {
                                            if (!list.$events.menuOpened) {
                                                app.$.each($scope.lists, (function (i, list) {
                                                    list.$events.menuOpened = false;
                                                }));
                                            }
                                            list.$events.menuOpened = !list.$events.menuOpened;
                                        }
                                    }
                                };
                                $scope.$watch("table.selectedItems", function (newValue, oldValue) {
                                    $scope.selection.commandBar.viewEnabled = newValue.length === 1;
                                    $scope.selection.commandBar.deleteEnabled = newValue.length > 0;
                                    $scope.selection.commandBar.settingsEnabled = newValue.length === 1;
                                    $scope.selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                                }, true);
                                $scope.$watch("selection.commandBar.searchTerm", function (newValue, oldValue) {
                                    $scope.table.rows.splice(0, $scope.table.rows.length);
                                    app.delay(function () {
                                        $scope.$apply(function () {
                                            var lists;
                                            if (newValue && newValue !== oldValue) {
                                                lists = [];
                                                app.$.each(factory.lists, function (i, list) {
                                                    if (list.$data && new RegExp(newValue, "i").test(list.$data.Title)) {
                                                        lists.push(list);
                                                    }
                                                });
                                            }
                                            else {
                                                lists = factory.lists;
                                            }
                                            $scope.lists = app.$angular.copy(lists);
                                        });
                                    }, options.delay);
                                }, false);
                            }]);
                        return deferred.promise();
                    };
                    return ListsView;
                }(app.App.Module.ListsViewBase));
                Module.ListsView = ListsView;
            })(Module = App.Module || (App.Module = {}));
        })(App = Angular.App || (Angular.App = {}));
    })(Angular || (Angular = {}));
    return new Angular.App();
});
