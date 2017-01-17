/// <reference path="typings/jquery/jquery.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
/// <reference path="typings/microsoft-ajax/microsoft.ajax.d.ts" />
define(["require", "exports", "pnp"], function (require, exports, $pnp) {
    "use strict";
    "use strict";
    var App;
    (function (App) {
        var AppBase = (function () {
            function AppBase() {
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
            AppBase.prototype.init = function (preloadedScripts) {
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
                }
                this.hostWebUrl = window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.webAbsoluteUrl) ? window._spPageContextInfo.webAbsoluteUrl : $pnp.util.getUrlParamByName("SPHostUrl");
                if ($pnp.util.stringIsNullOrEmpty(this.hostWebUrl)) {
                    throw "SPHostUrl url parameter must be specified!";
                }
                this.appWebUrl = window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.appWebUrl) ? window._spPageContextInfo.appWebUrl : $pnp.util.getUrlParamByName("SPAppWebUrl");
                if ($pnp.util.stringIsNullOrEmpty(this.appWebUrl)) {
                    throw "SPAppWebUrl url parameter must be specified!";
                }
                this.scriptBase = $pnp.util.combinePaths(this.hostWebUrl, window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.layoutsUrl) ? window._spPageContextInfo.layoutsUrl : "_layouts/15");
                this._initialized = true;
                self.$(self).trigger("app-init");
            };
            AppBase.prototype.ensureScript = function (url) {
                if (url) {
                    url = url.toLowerCase().replace("~sphost", this.hostWebUrl)
                        .replace("~spapp", this.appWebUrl)
                        .replace("~splayouts", this.scriptBase);
                    var scriptPromise = this._scriptPromises[url];
                    if (!scriptPromise) {
                        scriptPromise = this.$.cachedScript(url);
                        this._scriptPromises[url] = scriptPromise;
                    }
                    return scriptPromise;
                }
                return null;
            };
            AppBase.prototype.render = function (modules) {
                var self = this;
                if (!self._initialized) {
                    throw "App is not initialized!";
                }
                self.ensureScript("~splayouts/MicrosoftAjax.js").then(function () {
                    self.ensureScript("~splayouts/SP.Runtime.js").then(function () {
                        self.ensureScript("~splayouts/SP.RequestExecutor.js").then(function () {
                            self.ensureScript("~splayouts/SP.js").then(function () {
                                if ($pnp.util.isArray(modules)) {
                                    self.$.each(modules, function (i, module) {
                                        module.render();
                                    });
                                }
                                self.$(self).trigger("app-render");
                            });
                        });
                    });
                });
            };
            AppBase.prototype.is_initialized = function () {
                var self = this;
                return self._initialized;
            };
            AppBase.prototype.get_hostWebUrl = function () {
                var self = this;
                return self.hostWebUrl;
            };
            AppBase.prototype.get_appWebUrl = function () {
                var self = this;
                return self.appWebUrl;
            };
            AppBase.prototype.get_BasePermissions = function (permMask) {
                var permissions = new SP.BasePermissions();
                if (permMask) {
                    var permMaskHigh = permMask.length <= 10 ? 0 : parseInt(permMask.substring(2, permMask.length - 8), 16);
                    var permMaskLow = permMask.length <= 10 ? parseInt(permMask) : parseInt(permMask.substring(permMask.length - 8, permMask.length), 16);
                    permissions.initPropertiesFromJson({ "High": permMaskHigh, "Low": permMaskLow });
                }
                return permissions;
            };
            AppBase.prototype.getQueryParam = function (url, name) {
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
            return AppBase;
        }());
        App.AppBase = AppBase;
        var Module;
        (function (Module) {
            (function (RenderMethod) {
                RenderMethod[RenderMethod["Default"] = 0] = "Default";
                RenderMethod[RenderMethod["RenderListDataAsStream"] = 1] = "RenderListDataAsStream";
                RenderMethod[RenderMethod["RenderListFilterData"] = 2] = "RenderListFilterData";
                RenderMethod[RenderMethod["RenderListData"] = 3] = "RenderListData";
                RenderMethod[RenderMethod["GetItems"] = 4] = "GetItems";
            })(Module.RenderMethod || (Module.RenderMethod = {}));
            var RenderMethod = Module.RenderMethod;
            var ListViewBase = (function () {
                function ListViewBase(app, options) {
                    if (!app) {
                        throw "App must be specified for ListView!";
                    }
                    this._app = app;
                    this._options = options;
                }
                ListViewBase.prototype.addToken = function (query, token) {
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
                ListViewBase.prototype.getListItems = function (token /*, prevItemId?: number, pageLastRow?: number*/) {
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
                                    if (self._app.$.isArray(self._options.filters)) {
                                        for (var i = 0; i < self._options.filters.length; i++) {
                                            var filter = self._options.filters[i];
                                            if (filter) {
                                                query.query.add("FilterField" + (i + 1), filter.field);
                                                query.query.add("FilterValue" + (i + 1), filter.value);
                                            }
                                        }
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
                                    error: function (data, errorCode, errorMessage) {
                                        if (data.body) {
                                            try {
                                                var error = JSON.parse(data.body);
                                                if (error && error.error) {
                                                    errorMessage = error.error.message.value;
                                                }
                                            }
                                            catch (e) { }
                                        }
                                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                                        deferred.reject(data, errorCode, errorMessage);
                                    }
                                });
                            }
                            else {
                                self._app.$(self._app).trigger("app-error", ["List is not specified."]);
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
                                    error: function (data, errorCode, errorMessage) {
                                        if (data.body) {
                                            try {
                                                var error = JSON.parse(data.body);
                                                if (error && error.error) {
                                                    errorMessage = error.error.message.value;
                                                }
                                            }
                                            catch (e) { }
                                        }
                                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                                        deferred.reject(data, errorCode, errorMessage);
                                    }
                                });
                            }
                            else {
                                self._app.$(self._app).trigger("app-error", ["List is not specified."]);
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
                                    error: function (data, errorCode, errorMessage) {
                                        if (data.body) {
                                            try {
                                                var error = JSON.parse(data.body);
                                                if (error && error.error) {
                                                    errorMessage = error.error.message.value;
                                                }
                                            }
                                            catch (e) { }
                                        }
                                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                                        deferred.reject(data, errorCode, errorMessage);
                                    }
                                });
                            }
                            else {
                                self._app.$(self._app).trigger("app-error", ["List is not specified."]);
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
                                    error: function (data, errorCode, errorMessage) {
                                        if (data.body) {
                                            try {
                                                var error = JSON.parse(data.body);
                                                if (error && error.error) {
                                                    errorMessage = error.error.message.value;
                                                }
                                            }
                                            catch (e) { }
                                        }
                                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                                        deferred.reject(data, errorCode, errorMessage);
                                    }
                                });
                            }
                            else {
                                self._app.$(self._app).trigger("app-error", ["List is not specified."]);
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
                                }, function (sender, args) {
                                    self._app.$(self._app).trigger("app-error", [args.get_message(), args.get_stackTrace()]);
                                    deferred.reject(sender, args);
                                });
                            }
                            else {
                                self._app.$(self._app).trigger("app-error", ["List is not specified."]);
                            }
                            break;
                    }
                    return deferred.promise();
                };
                ListViewBase.prototype.get_app = function () {
                    return this._app;
                };
                ListViewBase.prototype.get_options = function () {
                    return this._options;
                };
                ListViewBase.prototype.render = function () {
                    this._app.$(self).trigger("model-render");
                };
                return ListViewBase;
            }());
            Module.ListViewBase = ListViewBase;
            var ListsViewBase = (function () {
                function ListsViewBase(app, options) {
                    if (!app) {
                        throw "App must be specified for ListView!";
                    }
                    this._app = app;
                    this._options = $pnp.util.extend(options, { delay: 1000 });
                }
                ListsViewBase.prototype.getLists = function () {
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
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
                                    if (error && error.message) {
                                        errorMessage = error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                    return deferred.promise();
                };
                ListsViewBase.prototype.getList = function (listId) {
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
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                    return deferred.promise();
                };
                ListsViewBase.prototype.updateList = function (listId, properties, digestValue) {
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
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                    return deferred.promise();
                };
                ListsViewBase.prototype.get_options = function () {
                    return this._options;
                };
                ListsViewBase.prototype.get_app = function () {
                    return this._app;
                };
                ListsViewBase.prototype.render = function () {
                    this._app.$(self).trigger("model-render");
                };
                return ListsViewBase;
            }());
            Module.ListsViewBase = ListsViewBase;
        })(Module = App.Module || (App.Module = {}));
    })(App = exports.App || (exports.App = {}));
});
