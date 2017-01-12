/// <reference path="typings/angularjs/angular.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
/// <reference path="typings/microsoft-ajax/microsoft.ajax.d.ts" />
define(["require", "exports", "pnp", "jquery"], function (require, exports, $pnp, $) {
    "use strict";
    //import * as $angular from "angular";
    var App = (function () {
        function App() {
            this.delay = (function () {
                var timer = 0;
                return function (callback, ms) {
                    clearTimeout(timer);
                    timer = setTimeout(callback, ms);
                };
            })();
            this._initialized = false;
            this._scriptPromises = {};
        }
        App.prototype.init = function (preloadedScripts) {
            var self = this;
            if (preloadedScripts) {
                var $_1 = preloadedScripts["jquery"];
                self.$ = $_1;
                if (self.$) {
                    self.$.cachedScript = function (url, options) {
                        options = self.$.extend(options || {}, {
                            dataType: "script",
                            cache: true,
                            url: url
                        });
                        return self.$.ajax(options);
                    };
                }
                else {
                    throw "jQuery is not loaded!";
                }
                var $angular = preloadedScripts["angular"];
                self.$angular = $angular;
                if (!self.$angular) {
                    throw "Angular is not loaded!";
                }
            }
            this.hostWebUrl = window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.webAbsoluteUrl) ? window._spPageContextInfo.webAbsoluteUrl : $pnp.util.getUrlParamByName("SPHostUrl");
            if ($pnp.util.stringIsNullOrEmpty(this.hostWebUrl)) {
                throw "SPHostUrl url parameter must be specified!";
            }
            this.appWebUrl = window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.appWebUrl) ? window._spPageContextInfo.appWebUrl : $pnp.util.getUrlParamByName("SPAppWebUrl");
            if ($pnp.util.stringIsNullOrEmpty(this.appWebUrl)) {
                throw "SPAppWebUrl url parameter must be specified!";
            }
            this.scriptBase = $pnp.util.combinePaths(this.hostWebUrl, "_layouts/15");
            this.spApp = this.$angular.module(App.SharePointAppName, [
                //'ngSanitize',
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
            }).filter('unsafe', function ($sce) {
                return $sce.trustAsHtml;
            }).directive('compile', [
                '$compile', function ($compile) {
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
            this._initialized = true;
            self.$(self).trigger("app-init");
        };
        App.prototype.ensureScript = function (url) {
            if (url) {
                url = url.toLowerCase().replace("~sphost", this.scriptBase);
                var scriptPromise = this._scriptPromises[url];
                if (!scriptPromise) {
                    scriptPromise = this.$.cachedScript(url);
                    this._scriptPromises[url] = scriptPromise;
                }
                return scriptPromise;
            }
            return null;
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
            self.ensureScript("~sphost/MicrosoftAjax.js").then(function () {
                self.ensureScript("~sphost/SP.Runtime.js").then(function () {
                    self.ensureScript("~sphost/SP.RequestExecutor.js").then(function () {
                        self.ensureScript("~sphost/SP.js").then(function () {
                            if ($pnp.util.isArray(modules)) {
                                self.$.each(modules, function (i, module) {
                                    module.render();
                                });
                            }
                            self.$angular.element(function () {
                                self.$angular.bootstrap(document, [App.SharePointAppName]);
                            });
                            self.$(self).trigger("app-render");
                        });
                    });
                });
            });
        };
        App.prototype.get_ListView = function (options) {
            var self = this;
            return new App.Module.ListView(self, options);
        };
        App.prototype.get_ListsView = function (options) {
            var self = this;
            return new App.Module.ListsView(self, options);
        };
        App.prototype.get_BasePermissions = function (permMask) {
            var permissions = new SP.BasePermissions();
            if (permMask) {
                var permMaskHigh = permMask.length <= 10 ? 0 : parseInt(permMask.substring(2, permMask.length - 8), 16);
                var permMaskLow = permMask.length <= 10 ? parseInt(permMask) : parseInt(permMask.substring(permMask.length - 8, permMask.length), 16);
                permissions.initPropertiesFromJson({ "High": permMaskHigh, "Low": permMaskLow });
            }
            return permissions;
        };
        App.prototype.getQueryParam = function (url, name) {
            var self = this;
            if (url) {
                var search = null;
                if (URL) {
                    search = new URL(url).search;
                }
                else {
                    var a = document.createElement('a');
                    a.href = url;
                    search = a.search;
                }
                if (search) {
                    while (search.startsWith("?") || search.startsWith("&")) {
                        search = search.slice(1, search.length);
                    }
                    var qParameters = search.split("&");
                    for (var i in qParameters) {
                        var qParameter = qParameters[i].split("=");
                        var key = decodeURIComponent(self.$(qParameter).get(0));
                        if (key && key.toUpperCase() === name.toUpperCase()) {
                            var value = decodeURIComponent(self.$(qParameter).get(1));
                            return value;
                        }
                    }
                }
            }
            return "";
        };
        App.SharePointAppName = "SharePointApp";
        App.SPServiceName = "SPService";
        return App;
    }());
    var App;
    (function (App) {
        var Module;
        (function (Module) {
            (function (RenderMethod) {
                RenderMethod[RenderMethod["Default"] = 0] = "Default";
                RenderMethod[RenderMethod["RenderListDataAsStream"] = 1] = "RenderListDataAsStream";
                RenderMethod[RenderMethod["RenderListData"] = 2] = "RenderListData";
                RenderMethod[RenderMethod["GetItems"] = 3] = "GetItems";
            })(Module.RenderMethod || (Module.RenderMethod = {}));
            var RenderMethod = Module.RenderMethod;
            var ListView = (function () {
                function ListView(app, options) {
                    if (!app) {
                        throw "App must be specified for ListView!";
                    }
                    this._app = app;
                    this._options = options;
                }
                ListView.prototype.getEntity = function (listItem) {
                    var permMask = listItem["PermMask"];
                    var permissions = this._app.get_BasePermissions(permMask);
                    var $permissions = {
                        edit: permissions.has(SP.PermissionKind.editListItems),
                        delete: permissions.has(SP.PermissionKind.deleteListItems)
                    };
                    var $events = { menuOpened: false, isSelected: false };
                    return { $data: listItem, $events: $events, $permissions: $permissions };
                };
                ListView.prototype.addToken = function (query, token) {
                    var self = this;
                    if (token) {
                        while (token.startsWith("?") || token.startsWith("&")) {
                            token = token.slice(1, token.length);
                        }
                        var qParameters = token.split("&");
                        for (var i in qParameters) {
                            var qParameter = qParameters[i].split("=");
                            var key = self._app.$(qParameter).get(0);
                            var value = self._app.$(qParameter).get(1);
                            if (key && value) {
                                query.query.add(key, value);
                            }
                        }
                    }
                };
                ListView.prototype.getListItems = function (token /*, prevItemId?: number, pageLastRow?: number*/) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    var url = null;
                    switch (self._options.renderMethod) {
                        case RenderMethod.RenderListDataAsStream:
                            var query = null;
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists;
                                query.concat("/GetById(@list)");
                                query.query.add("@list", "'" + self._options.listId + "'");
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                            }
                            if (query) {
                                query.concat("/RenderListDataAsStream");
                                if (!$pnp.util.stringIsNullOrEmpty(token)) {
                                    self.addToken(query, token);
                                }
                                else {
                                    if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                                        query.query.add("View", self._options.viewId);
                                    }
                                    if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                                        query.query.add("SortField", self._options.orderBy);
                                    }
                                    if (!$pnp.util.stringIsNullOrEmpty(self._options.sortAsc)) {
                                        query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                                    }
                                }
                                //if (!$pnp.util.stringIsNullOrEmpty(<any>prevItemId)) {
                                //    query.query.add("p_ID", prevItemId);
                                //}
                                //if (!$pnp.util.stringIsNullOrEmpty(<any>pageLastRow)) {
                                //    query.query.add("PageLastRow", pageLastRow);
                                //}
                                url = query.toUrlAndQuery();
                                var parameters = { "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": self._options.renderOptions };
                                if (!$pnp.util.stringIsNullOrEmpty(self._options.viewXml)) {
                                    parameters.ViewXml = self._options.viewXml;
                                }
                                if (!$pnp.util.stringIsNullOrEmpty(self._options.paged)) {
                                }
                                if (!$pnp.util.stringIsNullOrEmpty(self._options.rootFolder)) {
                                    parameters.FolderServerRelativeUrl = self._options.rootFolder;
                                }
                                var postBody = JSON.stringify({ "parameters": parameters });
                                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                                executor.executeAsync({
                                    url: url,
                                    method: "POST",
                                    headers: {
                                        "accept": "application/json;odata=verbose",
                                        "content-Type": "application/json;odata=verbose"
                                    },
                                    body: postBody,
                                    success: function (data) {
                                        var result = JSON.parse(data.body);
                                        deferred.resolve(result);
                                    },
                                    error: function (error) {
                                        deferred.reject(error);
                                    }
                                });
                            }
                            else {
                                deferred.reject("List is not specified.");
                            }
                            break;
                        case RenderMethod.RenderListData:
                            var query = null;
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                            }
                            if (query) {
                                query.concat("/renderListData(@viewXml)");
                                query.query.add("@viewXml", "'" + self._options.viewXml + "'");
                                if (!$pnp.util.stringIsNullOrEmpty(token)) {
                                    self.addToken(query, token);
                                }
                                else {
                                    if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                                        query.query.add("View", self._options.viewId);
                                    }
                                    if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                                        query.query.add("SortField", self._options.orderBy);
                                    }
                                    if (!$pnp.util.stringIsNullOrEmpty(self._options.sortAsc)) {
                                        query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                                    }
                                }
                                //if (!$pnp.util.stringIsNullOrEmpty(<any>prevItemId)) {
                                //    query.query.add("p_ID", prevItemId);
                                //}
                                //if (!$pnp.util.stringIsNullOrEmpty(<any>pageLastRow)) {
                                //    query.query.add("PageLastRow", pageLastRow);
                                //}
                                url = query.toUrlAndQuery();
                                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                                executor.executeAsync({
                                    url: url,
                                    method: "POST",
                                    headers: {
                                        "accept": "application/json;odata=verbose",
                                        "content-Type": "application/json;odata=verbose"
                                    },
                                    success: function (data) {
                                        var result = JSON.parse(JSON.parse(data.body).d.RenderListData);
                                        deferred.resolve({ ListData: result });
                                    },
                                    error: function (error) {
                                        deferred.reject(error);
                                    }
                                });
                            }
                            else {
                                deferred.reject("List is not specified.");
                            }
                            break;
                        case RenderMethod.GetItems:
                            var query = null;
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                            }
                            if (query) {
                                query.concat("/GetItems");
                                if ($pnp.util.isArray(self._options.expands)) {
                                    query = query.expand(self._options.expands);
                                }
                                url = query.toUrlAndQuery();
                                var postBody = JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": self._options.viewXml } });
                                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                                executor.executeAsync({
                                    url: url,
                                    method: "POST",
                                    body: postBody,
                                    headers: {
                                        "accept": "application/json;odata=verbose",
                                        "content-Type": "application/json;odata=verbose"
                                    },
                                    success: function (data) {
                                        var d = JSON.parse(data.body).d;
                                        var listData = { Row: d.results, NextHref: null, PrevHref: null };
                                        deferred.resolve({ ListData: listData });
                                    },
                                    error: function (error) {
                                        deferred.reject(error);
                                    }
                                });
                            }
                            else {
                                deferred.reject("List is not specified.");
                            }
                            break;
                        case RenderMethod.Default:
                            var query = null;
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId).items;
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl).items;
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                                query = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listTitle).items;
                            }
                            if (query) {
                                if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                                    query = query.orderBy(self._options.orderBy, self._options.sortAsc);
                                }
                                if (!$pnp.util.stringIsNullOrEmpty(self._options.filter)) {
                                    query = query.filter(self._options.filter);
                                }
                                if (self._options.limit > 0) {
                                    query = query.top(self._options.limit);
                                }
                                if ($pnp.util.isArray(self._options.expands)) {
                                    query = query.expand(self._options.expands);
                                }
                                if (!$pnp.util.stringIsNullOrEmpty(token)) {
                                    query.query.add("$skiptoken", encodeURIComponent(token));
                                }
                                url = query.toUrlAndQuery();
                                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                                executor.executeAsync({
                                    url: url,
                                    method: "GET",
                                    headers: {
                                        "accept": "application/json;odata=verbose",
                                        "content-Type": "application/json;odata=verbose"
                                    },
                                    success: function (data) {
                                        var d = JSON.parse(data.body).d;
                                        var listData = { Row: d.results, NextHref: null, PrevHref: null };
                                        listData.NextHref = self._app.getQueryParam(d["__next"], "$skiptoken");
                                        listData.PrevHref = self._app.getQueryParam(d["__prev"], "$skiptoken");
                                        deferred.resolve({ ListData: listData });
                                    },
                                    error: function (error) {
                                        deferred.reject(error);
                                    }
                                });
                            }
                            else {
                                deferred.reject("List is not specified.");
                            }
                            break;
                        default:
                            var context = new SP.ClientContext(self._app.appWebUrl);
                            var factory = new SP.ProxyWebRequestExecutorFactory(self._app.appWebUrl);
                            context.set_webRequestExecutorFactory(factory);
                            var appContextSite = new SP.AppContextSite(context, self._app.hostWebUrl);
                            var web = appContextSite.get_web();
                            var list = null;
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                                list = web.get_lists().getById(self._options.listId);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                                list = web.getList(self._options.listUrl);
                            }
                            else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                                list = web.get_lists().getByTitle(self._options.listTitle);
                            }
                            if (list) {
                                var camlQuery = new SP.CamlQuery();
                                camlQuery.set_viewXml(self._options.viewXml);
                                if (!$pnp.util.stringIsNullOrEmpty(token)) {
                                    var position = new SP.ListItemCollectionPosition();
                                    position.set_pagingInfo(token);
                                    camlQuery.set_listItemCollectionPosition(position);
                                }
                                var items = list.getItems(camlQuery);
                                context.load(items);
                                context.executeQueryAsync(function () {
                                    var listData = { Row: self._app.$.map(items.get_data(), function (item) { return item.get_fieldValues(); }), NextHref: null, PrevHref: null };
                                    var position = items.get_listItemCollectionPosition();
                                    if (position) {
                                        listData.NextHref = position.get_pagingInfo();
                                    }
                                    deferred.resolve({ ListData: listData });
                                }, deferred.reject);
                            }
                            else {
                                deferred.reject("List is not specified.");
                            }
                            break;
                    }
                    return deferred.promise();
                };
                ListView.prototype.render = function () {
                    var self = this;
                    var allTokens = [];
                    self._app.spApp.factory("ListViewFactory", function ($q, $http) {
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
                                    self._app.$.each(data.ListData.Row, (function (i, listItem) {
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
                    var deferred = self._app.$.Deferred();
                    self._app.spApp.controller(self._options.controllerName, [
                        '$scope', 'ListViewFactory', App.SPServiceName, function ($scope, factory, service) {
                            $scope.loading = true;
                            $scope.listItems = [];
                            factory.getListItems().then(function () {
                                $scope.listItems = self._app.$angular.copy(factory.listItems);
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
                                            listItem = self._app.$(selectedItems).get(0);
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
                                            self._app.$.each($scope.table.rows, function (i, item) {
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
                                            self._app.$.each($scope.listItems, (function (i, listItem) {
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
                                        var token = self._options.appendRows === true ? null : factory.getToken(0);
                                        $scope.table.rows.splice(0, $scope.table.rows.length);
                                        factory.getListItems(token).then(function () {
                                            //(<any>$scope).selection.commandBar.clearSelection();
                                            if ($pnp.util.stringIsNullOrEmpty(token) || !self._options.appendRows) {
                                                $scope.listItems = self._app.$angular.copy(factory.listItems);
                                            }
                                            else {
                                                $scope.listItems = self._app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                            }
                                            $scope.selection.pager.first = self._options.appendRows === true ? ($scope.listItems.length > 0 ? 1 : 0) : factory.$first;
                                            $scope.selection.pager.last = self._options.appendRows === true ? ($scope.listItems.length) : factory.$last;
                                            $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                            $scope.selection.pager.prevEnabled = self._options.appendRows !== true && !$pnp.util.stringIsNullOrEmpty(factory.getToken(-1));
                                            if (!token) {
                                                allTokens = [];
                                            }
                                            else {
                                                allTokens.pop();
                                            }
                                            allTokens.push(factory.$nextToken);
                                            //(<any>$scope).loading = false;
                                            deferred.resolve();
                                        }, deferred.reject);
                                    },
                                    next: function (offset) {
                                        if (offset === void 0) { offset = 1; }
                                        if (!$scope.selection.pager.nextEnabled)
                                            return;
                                        $scope.table.rows.splice(0, $scope.table.rows.length);
                                        var token = factory.getToken(offset);
                                        factory.getListItems(token).then(function () {
                                            //(<any>$scope).selection.commandBar.clearSelection();
                                            if ($pnp.util.stringIsNullOrEmpty(token) || !self._options.appendRows) {
                                                $scope.listItems = self._app.$angular.copy(factory.listItems);
                                            }
                                            else {
                                                $scope.listItems = self._app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                            }
                                            $scope.selection.pager.first = self._options.appendRows === true ? ($scope.listItems.length > 0 ? 1 : 0) : factory.$first;
                                            $scope.selection.pager.last = self._options.appendRows === true ? ($scope.listItems.length) : factory.$last;
                                            $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                            $scope.selection.pager.prevEnabled = self._options.appendRows !== true && !$pnp.util.stringIsNullOrEmpty(factory.getToken(-1));
                                            if ($pnp.util.stringIsNullOrEmpty(token)) {
                                                allTokens = [];
                                            }
                                            allTokens.push(factory.$nextToken);
                                            //(<any>$scope).loading = false;
                                            deferred.resolve();
                                        }, deferred.reject);
                                    },
                                    prev: function (offset) {
                                        if (offset === void 0) { offset = 1; }
                                        if (!$scope.selection.pager.prevEnabled)
                                            return;
                                        $scope.table.rows.splice(0, $scope.table.rows.length);
                                        offset = Math.min(-1, -offset);
                                        var token = self._options.appendRows === true ? null : factory.getToken(offset);
                                        factory.getListItems(token).then(function () {
                                            //(<any>$scope).selection.commandBar.clearSelection();
                                            if ($pnp.util.stringIsNullOrEmpty(token) || !self._options.appendRows) {
                                                $scope.listItems = self._app.$angular.copy(factory.listItems);
                                            }
                                            else {
                                                $scope.listItems = self._app.$angular.copy([].concat($scope.listItems).concat(factory.listItems));
                                            }
                                            $scope.selection.pager.first = self._options.appendRows === true ? ($scope.listItems.length > 0 ? 1 : 0) : factory.$first;
                                            $scope.selection.pager.last = self._options.appendRows === true ? ($scope.listItems.length) : factory.$last;
                                            $scope.selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                            $scope.selection.pager.prevEnabled = self._options.appendRows !== true && !$pnp.util.stringIsNullOrEmpty(factory.getToken(-1));
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
                                            //(<any>$scope).loading = false;
                                            deferred.resolve();
                                        }, deferred.reject);
                                    },
                                }
                            };
                            $scope.$watch('table.selectedItems', function (newValue, oldValue) {
                                $scope.selection.commandBar.viewEnabled = newValue.length === 1;
                                $scope.selection.commandBar.deleteEnabled = newValue.length > 0;
                                $scope.selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                            }, true);
                            $scope.$watch('selection.commandBar.searchTerm', function (newValue, oldValue) {
                                //(<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                                //self._app.delay(() => {
                                //    $scope.$apply(function () {
                                //    });
                                //}, self._options.delay);
                            }, false);
                            if (typeof self._options.onload === "function") {
                                self._options.onload($scope, factory);
                            }
                        }
                    ]);
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
                    this._options = this._app.$.extend(true, { delay: 1000 }, options);
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
                            $scope.lists = [];
                            factory.getLists().then(function () {
                                $scope.lists = self._app.$angular.copy(factory.lists);
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
                                            $scope.selection.settings.data = self._app.$.extend(true, {}, data);
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
                                            list = self._app.$(selectedItems).get(0);
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
                                            list = self._app.$(selectedItems).get(0);
                                        }
                                        if (list) {
                                            window.location.href = "/Home/List?ListId=" + list.$data.Id + "&SPHostUrl=" + decodeURIComponent(self._app.hostWebUrl) + "&SPAppWebUrl=" + decodeURIComponent(self._app.appWebUrl);
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
                                            self._app.$.each($scope.table.rows, function (i, item) {
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
                                            self._app.$.each($scope.lists, (function (i, list) {
                                                list.$events.menuOpened = false;
                                            }));
                                        }
                                        list.$events.menuOpened = !list.$events.menuOpened;
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
                                $scope.table.rows.splice(0, $scope.table.rows.length);
                                self._app.delay(function () {
                                    $scope.$apply(function () {
                                        var lists;
                                        if (newValue && newValue !== oldValue) {
                                            lists = [];
                                            self._app.$.each(factory.lists, function (i, list) {
                                                if (list.$data && new RegExp(newValue, 'i').test(list.$data.Title)) {
                                                    lists.push(list);
                                                }
                                            });
                                        }
                                        else {
                                            lists = factory.lists;
                                        }
                                        $scope.lists = self._app.$angular.copy(lists);
                                    });
                                }, self._options.delay);
                            }, false);
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
