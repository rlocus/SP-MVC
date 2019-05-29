var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
define(["require", "exports", "pnp", "jquery", "./app.module"], function (require, exports, $pnp, $, app) {
    "use strict";
    var Angular;
    (function (Angular) {
        "use strict";
        var App = (function (_super) {
            __extends(App, _super);
            function App() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            App.prototype.init = function (preloadedScripts) {
                var self = this;
                var $angular = preloadedScripts["angular"];
                self.$angular = $angular;
                if (!self.$angular) {
                    throw "angular is not loaded!";
                }
                self.ngModule = self.$angular.module(App.SharePointAppName, [
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
                                return scope.$eval(attrs.compile);
                            }, function (value) {
                                element.html(value);
                                $compile(element.contents())(scope);
                            });
                        };
                    }
                ]).directive('ngAppFrame', ['$timeout', '$window', function ($timeout, $window) {
                        return {
                            restrict: 'E',
                            link: function (scope, element, attrs) {
                                scope.$on('resizeframe', function () {
                                    $timeout(function () {
                                        var contentHeight = element[0].offsetParent.clientHeight;
                                        var resizeMessage = '<message senderId={Sender_ID}>resize({Width}, {Height}px)</message>';
                                        var senderId = $pnp.util.getUrlParamByName("SenderId").split("#")[0];
                                        var step = 30, finalHeight;
                                        finalHeight = (step - (contentHeight % step)) + contentHeight;
                                        resizeMessage = resizeMessage.replace("{Sender_ID}", senderId);
                                        resizeMessage = resizeMessage.replace("{Height}", finalHeight);
                                        resizeMessage = resizeMessage.replace("{Width}", "100%");
                                        window.parent.postMessage(resizeMessage, "*");
                                    }, 0, false);
                                });
                            }
                        };
                    }]);
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
                return new App.Module.ListView(self, options);
            };
            App.prototype.get_ListsView = function (options) {
                var self = this;
                return new App.Module.ListsView(self, options);
            };
            App.SharePointAppName = "SharePointApp";
            App.SPServiceName = "SPService";
            return App;
        }(app.App.AppBase));
        Angular.App = App;
        (function (App) {
            var Module;
            (function (Module) {
                var ListView = (function (_super) {
                    __extends(ListView, _super);
                    function ListView(app, options) {
                        var _this = this;
                        if (!options.delay) {
                            options.delay = 1000;
                        }
                        _this = _super.call(this, app, options) || this;
                        return _this;
                    }
                    ListView.prototype.getEntity = function (listItem) {
                        var permMask = listItem["PermMask"];
                        var permissions = this.get_app().get_BasePermissions(permMask);
                        var $permissions = {
                            edit: permissions.has(SP.PermissionKind.editListItems),
                            "delete": permissions.has(SP.PermissionKind.deleteListItems)
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
                                    factory.listItems = [];
                                    if (data.ListData) {
                                        app.$.each(data.ListData.Row, (function (i, listItem) {
                                            factory.listItems.push(self.getEntity(listItem));
                                        }));
                                        factory.$nextToken = data.ListData.NextHref;
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
                        app.ngModule.controller(options.controllerName, [
                            "$scope", "ListViewFactory", App.SPServiceName, function ($scope, factory, service) {
                                $scope.loading = true;
                                $scope.listItems = [];
                                $scope.$parent.$broadcast('resizeframe');
                                factory.getListItems().then(function () {
                                    allTokens.push(factory.$nextToken);
                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                    $scope.selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                    $scope.selection.pager.last = factory.$last > 0 ? factory.$last : ($scope.selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                    $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                    $scope.loading = false;
                                    $scope.$parent.$broadcast('resizeframe');
                                });
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
                                        "delete": function (listItem) {
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
                                            selectedItems.splice(0, selectedItems.length);
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
                                                $scope.table.selectedItems.splice(0, $scope.table.selectedItems.length);
                                                if (!token) {
                                                    allTokens = [];
                                                }
                                                else {
                                                    allTokens.pop();
                                                }
                                                allTokens.push(factory.$nextToken);
                                                var prevToken = factory.getToken(-1);
                                                if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                }
                                                else {
                                                    $scope.listItems = app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                                }
                                                $scope.selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                                $scope.selection.pager.last = factory.$last > 0 ? factory.$last : ($scope.selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                                $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                $scope.selection.pager.prevEnabled = options.appendRows !== true && ($pnp.util.stringIsNullOrEmpty(factory.$prevToken) ? !$pnp.util.stringIsNullOrEmpty(factory.getToken(0)) : true);
                                                $scope.$parent.$broadcast('resizeframe');
                                            });
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
                                                $scope.table.selectedItems.splice(0, $scope.table.selectedItems.length);
                                                if ($pnp.util.stringIsNullOrEmpty(token)) {
                                                    allTokens = [];
                                                }
                                                allTokens.push(factory.$nextToken);
                                                var prevToken = factory.getToken(-1);
                                                if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                }
                                                else {
                                                    $scope.listItems = app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                                }
                                                $scope.selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                                $scope.selection.pager.last = factory.$last > 0 ? factory.$last : ($scope.selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                                $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                $scope.selection.pager.prevEnabled = options.appendRows !== true && ($pnp.util.stringIsNullOrEmpty(factory.$prevToken) ? !$pnp.util.stringIsNullOrEmpty(factory.getToken(0)) : true);
                                                $scope.$parent.$broadcast('resizeframe');
                                            });
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
                                                $scope.table.selectedItems.splice(0, $scope.table.selectedItems.length);
                                                if ($pnp.util.stringIsNullOrEmpty(token)) {
                                                    allTokens = [];
                                                }
                                                else {
                                                    var skipNext = 1 - offset;
                                                    while (skipNext > 0) {
                                                        allTokens.pop();
                                                        skipNext--;
                                                    }
                                                }
                                                allTokens.push(factory.$nextToken);
                                                var prevToken = factory.getToken(-1);
                                                if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                }
                                                else {
                                                    $scope.listItems = app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                                }
                                                $scope.selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                                $scope.selection.pager.last = factory.$last > 0 ? factory.$last : ($scope.selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                                $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                $scope.selection.pager.prevEnabled = options.appendRows !== true && ($pnp.util.stringIsNullOrEmpty(factory.$prevToken) ? !$pnp.util.stringIsNullOrEmpty(factory.getToken(0)) : true);
                                                $scope.$parent.$broadcast('resizeframe');
                                            });
                                        }
                                    }
                                };
                                $scope.$watch("table.selectedItems", function (newValue, oldValue) {
                                    $scope.selection.commandBar.viewEnabled = newValue.length === 1;
                                    $scope.selection.commandBar.deleteEnabled = newValue.length > 0;
                                    $scope.selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                                }, true);
                                $scope.$watch("selection.commandBar.searchTerm", function (newValue, oldValue) {
                                    $scope.table.rows.splice(0, $scope.table.rows.length);
                                    app.delay(function () {
                                        $scope.$apply(function () {
                                            var listItems;
                                            if (newValue !== oldValue) {
                                                options.queryBuilder.clear();
                                                if (newValue) {
                                                    var filters = new Array();
                                                    app.$(self).trigger("search-item", [filters, newValue, $scope, factory]);
                                                    options.queryBuilder.appendAndWithAny.apply(options.queryBuilder, filters);
                                                }
                                                $scope.selection.pager.prevEnabled = false;
                                                $scope.selection.pager.nextEnabled = false;
                                                factory.getListItems(null).then(function () {
                                                    allTokens = [];
                                                    allTokens.push(factory.$nextToken);
                                                    $scope.listItems = app.$angular.copy(factory.listItems);
                                                    $scope.selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                                    $scope.selection.pager.last = factory.$last > 0 ? factory.$last : ($scope.selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                                    $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                                    $scope.$parent.$broadcast('resizeframe');
                                                });
                                            }
                                        });
                                    }, options.delay);
                                }, false);
                                app.$(self).trigger("model-render", [$scope, factory]);
                            }
                        ]);
                    };
                    return ListView;
                }(app.App.Module.ListViewBase));
                Module.ListView = ListView;
                var ListsView = (function (_super) {
                    __extends(ListsView, _super);
                    function ListsView(app, options) {
                        var _this = this;
                        if (!options.delay) {
                            options.delay = 1000;
                        }
                        _this = _super.call(this, app, options) || this;
                        return _this;
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
                                        "delete": function (list) {
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
                                            selectedItems.splice(0, selectedItems.length);
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
                                    $scope.table.selectedItems.splice(0, $scope.table.selectedItems.length);
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
