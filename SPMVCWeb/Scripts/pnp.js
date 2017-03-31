/**
 * sp-pnp-js v2.0.2 - A JavaScript library for SharePoint development.
 * MIT (https://github.com/SharePoint/PnP-JS-Core/blob/master/LICENSE)
 * Copyright (c) 2017 Microsoft
 * docs: http://officedev.github.io/PnP-JS-Core
 * source: https://github.com/SharePoint/PnP-JS-Core
 * bugs: https://github.com/SharePoint/PnP-JS-Core/issues
 */
(function webpackUniversalModuleDefinition(root, factory) {
    if (typeof exports === 'object' && typeof module === 'object')
        module.exports = factory();
    else if (typeof define === 'function' && define.amd)
        define([], factory);
    else if (typeof exports === 'object')
        exports["$pnp"] = factory();
    else
        root["$pnp"] = factory();
})(this, function () {
    return /******/ (function (modules) { // webpackBootstrap
        /******/ 	// The module cache
        /******/ 	var installedModules = {};
        /******/
        /******/ 	// The require function
        /******/ 	function __webpack_require__(moduleId) {
            /******/
            /******/ 		// Check if module is in cache
            /******/ 		if (installedModules[moduleId])
                /******/ 			return installedModules[moduleId].exports;
            /******/
            /******/ 		// Create a new module (and put it into the cache)
            /******/ 		var module = installedModules[moduleId] = {
                /******/ 			i: moduleId,
                /******/ 			l: false,
                /******/ 			exports: {}
                /******/
            };
            /******/
            /******/ 		// Execute the module function
            /******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
            /******/
            /******/ 		// Flag the module as loaded
            /******/ 		module.l = true;
            /******/
            /******/ 		// Return the exports of the module
            /******/ 		return module.exports;
            /******/
        }
        /******/
        /******/
        /******/ 	// expose the modules object (__webpack_modules__)
        /******/ 	__webpack_require__.m = modules;
        /******/
        /******/ 	// expose the module cache
        /******/ 	__webpack_require__.c = installedModules;
        /******/
        /******/ 	// identity function for calling harmony imports with the correct context
        /******/ 	__webpack_require__.i = function (value) { return value; };
        /******/
        /******/ 	// define getter function for harmony exports
        /******/ 	__webpack_require__.d = function (exports, name, getter) {
            /******/ 		if (!__webpack_require__.o(exports, name)) {
                /******/ 			Object.defineProperty(exports, name, {
                    /******/ 				configurable: false,
                    /******/ 				enumerable: true,
                    /******/ 				get: getter
                    /******/
                });
                /******/
            }
            /******/
        };
        /******/
        /******/ 	// getDefaultExport function for compatibility with non-harmony modules
        /******/ 	__webpack_require__.n = function (module) {
            /******/ 		var getter = module && module.__esModule ?
            /******/ 			function getDefault() { return module['default']; } :
            /******/ 			function getModuleExports() { return module; };
            /******/ 		__webpack_require__.d(getter, 'a', getter);
            /******/ 		return getter;
            /******/
        };
        /******/
        /******/ 	// Object.prototype.hasOwnProperty.call
        /******/ 	__webpack_require__.o = function (object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
        /******/
        /******/ 	// __webpack_public_path__
        /******/ 	__webpack_require__.p = "/assets/";
        /******/
        /******/ 	// Load entry module and return exports
        /******/ 	return __webpack_require__(__webpack_require__.s = 49);
        /******/
    })
    /************************************************************************/
    /******/([
    /* 0 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";
        /* WEBPACK VAR INJECTION */(function (global) {
            var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
                var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
                if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
                else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
                return c > 3 && r && Object.defineProperty(target, key, r), r;
            };
            Object.defineProperty(exports, "__esModule", { value: true });
            var decorators_1 = __webpack_require__(47);
            var pnplibconfig_1 = __webpack_require__(4);
            var Util = (function () {
                function Util() {
                }
                /**
                 * Gets a callback function which will maintain context across async calls.
                 * Allows for the calling pattern getCtxCallback(thisobj, method, methodarg1, methodarg2, ...)
                 *
                 * @param context The object that will be the 'this' value in the callback
                 * @param method The method to which we will apply the context and parameters
                 * @param params Optional, additional arguments to supply to the wrapped method when it is invoked
                 */
                Util.getCtxCallback = function (context, method) {
                    var params = [];
                    for (var _i = 2; _i < arguments.length; _i++) {
                        params[_i - 2] = arguments[_i];
                    }
                    return function () {
                        method.apply(context, params);
                    };
                };
                /**
                 * Tests if a url param exists
                 *
                 * @param name The name of the url paramter to check
                 */
                Util.urlParamExists = function (name) {
                    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
                    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)");
                    return regex.test(location.search);
                };
                /**
                 * Gets a url param value by name
                 *
                 * @param name The name of the paramter for which we want the value
                 */
                Util.getUrlParamByName = function (name) {
                    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
                    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)");
                    var results = regex.exec(location.search);
                    return results == null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
                };
                /**
                 * Gets a url param by name and attempts to parse a bool value
                 *
                 * @param name The name of the paramter for which we want the boolean value
                 */
                Util.getUrlParamBoolByName = function (name) {
                    var p = this.getUrlParamByName(name);
                    var isFalse = (p === "" || /false|0/i.test(p));
                    return !isFalse;
                };
                /**
                 * Inserts the string s into the string target as the index specified by index
                 *
                 * @param target The string into which we will insert s
                 * @param index The location in target to insert s (zero based)
                 * @param s The string to insert into target at position index
                 */
                Util.stringInsert = function (target, index, s) {
                    if (index > 0) {
                        return target.substring(0, index) + s + target.substring(index, target.length);
                    }
                    return s + target;
                };
                /**
                 * Adds a value to a date
                 *
                 * @param date The date to which we will add units, done in local time
                 * @param interval The name of the interval to add, one of: ['year', 'quarter', 'month', 'week', 'day', 'hour', 'minute', 'second']
                 * @param units The amount to add to date of the given interval
                 *
                 * http://stackoverflow.com/questions/1197928/how-to-add-30-minutes-to-a-javascript-date-object
                 */
                Util.dateAdd = function (date, interval, units) {
                    var ret = new Date(date.toLocaleString()); // don't change original date
                    switch (interval.toLowerCase()) {
                        case "year":
                            ret.setFullYear(ret.getFullYear() + units);
                            break;
                        case "quarter":
                            ret.setMonth(ret.getMonth() + 3 * units);
                            break;
                        case "month":
                            ret.setMonth(ret.getMonth() + units);
                            break;
                        case "week":
                            ret.setDate(ret.getDate() + 7 * units);
                            break;
                        case "day":
                            ret.setDate(ret.getDate() + units);
                            break;
                        case "hour":
                            ret.setTime(ret.getTime() + units * 3600000);
                            break;
                        case "minute":
                            ret.setTime(ret.getTime() + units * 60000);
                            break;
                        case "second":
                            ret.setTime(ret.getTime() + units * 1000);
                            break;
                        default:
                            ret = undefined;
                            break;
                    }
                    return ret;
                };
                /**
                 * Loads a stylesheet into the current page
                 *
                 * @param path The url to the stylesheet
                 * @param avoidCache If true a value will be appended as a query string to avoid browser caching issues
                 */
                Util.loadStylesheet = function (path, avoidCache) {
                    if (avoidCache) {
                        path += "?" + encodeURIComponent((new Date()).getTime().toString());
                    }
                    var head = document.getElementsByTagName("head");
                    if (head.length > 0) {
                        var e = document.createElement("link");
                        head[0].appendChild(e);
                        e.setAttribute("type", "text/css");
                        e.setAttribute("rel", "stylesheet");
                        e.setAttribute("href", path);
                    }
                };
                /**
                 * Combines an arbitrary set of paths ensuring that the slashes are normalized
                 *
                 * @param paths 0 to n path parts to combine
                 */
                Util.combinePaths = function () {
                    var paths = [];
                    for (var _i = 0; _i < arguments.length; _i++) {
                        paths[_i] = arguments[_i];
                    }
                    return paths
                        .filter(function (path) { return typeof path !== "undefined" && path !== null; })
                        .map(function (path) { return path.replace(/^[\\|\/]/, "").replace(/[\\|\/]$/, ""); })
                        .join("/")
                        .replace(/\\/g, "/");
                };
                /**
                 * Gets a random string of chars length
                 *
                 * @param chars The length of the random string to generate
                 */
                Util.getRandomString = function (chars) {
                    var text = new Array(chars);
                    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
                    for (var i = 0; i < chars; i++) {
                        text[i] = possible.charAt(Math.floor(Math.random() * possible.length));
                    }
                    return text.join("");
                };
                /**
                 * Gets a random GUID value
                 *
                 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
                 */
                /* tslint:disable no-bitwise */
                Util.getGUID = function () {
                    var d = new Date().getTime();
                    var guid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
                        var r = (d + Math.random() * 16) % 16 | 0;
                        d = Math.floor(d / 16);
                        return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
                    });
                    return guid;
                };
                /* tslint:enable */
                /**
                 * Determines if a given value is a function
                 *
                 * @param candidateFunction The thing to test for being a function
                 */
                Util.isFunction = function (candidateFunction) {
                    return typeof candidateFunction === "function";
                };
                /**
                 * @returns whether the provided parameter is a JavaScript Array or not.
                */
                Util.isArray = function (array) {
                    if (Array.isArray) {
                        return Array.isArray(array);
                    }
                    return array && typeof array.length === "number" && array.constructor === Array;
                };
                /**
                 * Determines if a string is null or empty or undefined
                 *
                 * @param s The string to test
                 */
                Util.stringIsNullOrEmpty = function (s) {
                    return typeof s === "undefined" || s === null || s === "";
                };
                /**
                 * Provides functionality to extend the given object by doing a shallow copy
                 *
                 * @param target The object to which properties will be copied
                 * @param source The source object from which properties will be copied
                 * @param noOverwrite If true existing properties on the target are not overwritten from the source
                 *
                 */
                Util.extend = function (target, source, noOverwrite) {
                    if (noOverwrite === void 0) { noOverwrite = false; }
                    if (source === null || typeof source === "undefined") {
                        return target;
                    }
                    // ensure we don't overwrite things we don't want overwritten
                    var check = noOverwrite ? function (o, i) { return !(i in o); } : function () { return true; };
                    return Object.getOwnPropertyNames(source)
                        .filter(function (v) { return check(target, v); })
                        .reduce(function (t, v) {
                            t[v] = source[v];
                            return t;
                        }, target);
                };
                /**
                 * Determines if a given url is absolute
                 *
                 * @param url The url to check to see if it is absolute
                 */
                Util.isUrlAbsolute = function (url) {
                    return /^https?:\/\/|^\/\//i.test(url);
                };
                /**
                 * Attempts to make the supplied relative url absolute based on the _spPageContextInfo object, if available
                 *
                 * @param url The relative url to make absolute
                 */
                Util.makeUrlAbsolute = function (url) {
                    if (Util.isUrlAbsolute(url)) {
                        return url;
                    }
                    if (typeof global._spPageContextInfo !== "undefined") {
                        if (global._spPageContextInfo.hasOwnProperty("webAbsoluteUrl")) {
                            return Util.combinePaths(global._spPageContextInfo.webAbsoluteUrl, url);
                        }
                        else if (global._spPageContextInfo.hasOwnProperty("webServerRelativeUrl")) {
                            return Util.combinePaths(global._spPageContextInfo.webServerRelativeUrl, url);
                        }
                    }
                    else {
                        return url;
                    }
                };
                /**
                 * Ensures that a given url is absolute for the current web based on context
                 *
                 * @param candidateUrl The url to make absolute
                 *
                 */
                Util.toAbsoluteUrl = function (candidateUrl) {
                    return new Promise(function (resolve) {
                        if (Util.isUrlAbsolute(candidateUrl)) {
                            // if we are already absolute, then just return the url
                            return resolve(candidateUrl);
                        }
                        if (pnplibconfig_1.RuntimeConfig.baseUrl !== null) {
                            // base url specified either with baseUrl of spfxContext config property
                            return resolve(Util.combinePaths(pnplibconfig_1.RuntimeConfig.baseUrl, candidateUrl));
                        }
                        if (typeof global._spPageContextInfo !== "undefined") {
                            // operating in classic pages
                            if (global._spPageContextInfo.hasOwnProperty("webAbsoluteUrl")) {
                                return resolve(Util.combinePaths(global._spPageContextInfo.webAbsoluteUrl, candidateUrl));
                            }
                            else if (global._spPageContextInfo.hasOwnProperty("webServerRelativeUrl")) {
                                return resolve(Util.combinePaths(global._spPageContextInfo.webServerRelativeUrl, candidateUrl));
                            }
                        }
                        // does window.location exist and have _layouts in it?
                        if (typeof global.location !== "undefined") {
                            var index = global.location.toString().toLowerCase().indexOf("/_layouts/");
                            if (index > 0) {
                                // we are likely in the workbench in /_layouts/
                                return resolve(Util.combinePaths(global.location.toString().substr(0, index), candidateUrl));
                            }
                        }
                        return resolve(candidateUrl);
                    });
                };
                return Util;
            }());
            __decorate([
                decorators_1.deprecated("The Util.makeUrlAbsolute method is deprecated and will be removed from future releases. Use Util.toAbsoluteUrl instead")
            ], Util, "makeUrlAbsolute", null);
            exports.Util = Util;

            /* WEBPACK VAR INJECTION */
        }.call(exports, __webpack_require__(30)))

        /***/
    }),
    /* 1 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        var collections_1 = __webpack_require__(6);
        var odata_1 = __webpack_require__(3);
        var pnplibconfig_1 = __webpack_require__(4);
        var exceptions_1 = __webpack_require__(2);
        var queryablerequest_1 = __webpack_require__(42);
        /**
         * Queryable Base Class
         *
         */
        var Queryable = (function () {
            /**
             * Creates a new instance of the Queryable class
             *
             * @constructor
             * @param baseUrl A string or Queryable that should form the base part of the url
             *
             */
            function Queryable(baseUrl, path) {
                this._query = new collections_1.Dictionary();
                this._batch = null;
                if (typeof baseUrl === "string") {
                    // we need to do some extra parsing to get the parent url correct if we are
                    // being created from just a string.
                    var urlStr = baseUrl;
                    if (util_1.Util.isUrlAbsolute(urlStr) || urlStr.lastIndexOf("/") < 0) {
                        this._parentUrl = urlStr;
                        this._url = util_1.Util.combinePaths(urlStr, path);
                    }
                    else if (urlStr.lastIndexOf("/") > urlStr.lastIndexOf("(")) {
                        // .../items(19)/fields
                        var index = urlStr.lastIndexOf("/");
                        this._parentUrl = urlStr.slice(0, index);
                        path = util_1.Util.combinePaths(urlStr.slice(index), path);
                        this._url = util_1.Util.combinePaths(this._parentUrl, path);
                    }
                    else {
                        // .../items(19)
                        var index = urlStr.lastIndexOf("(");
                        this._parentUrl = urlStr.slice(0, index);
                        this._url = util_1.Util.combinePaths(urlStr, path);
                    }
                }
                else {
                    var q = baseUrl;
                    this._parentUrl = q._url;
                    var target = q._query.get("@target");
                    if (target !== null) {
                        this._query.add("@target", target);
                    }
                    this._url = util_1.Util.combinePaths(this._parentUrl, path);
                }
            }
            /**
             * Directly concatonates the supplied string to the current url, not normalizing "/" chars
             *
             * @param pathPart The string to concatonate to the url
             */
            Queryable.prototype.concat = function (pathPart) {
                this._url += pathPart;
            };
            /**
             * Appends the given string and normalizes "/" chars
             *
             * @param pathPart The string to append
             */
            Queryable.prototype.append = function (pathPart) {
                this._url = util_1.Util.combinePaths(this._url, pathPart);
            };
            /**
             * Blocks a batch call from occuring, MUST be cleared by calling the returned function
             */
            Queryable.prototype.addBatchDependency = function () {
                if (this.hasBatch) {
                    return this._batch.addDependency();
                }
                return function () { return null; };
            };
            Object.defineProperty(Queryable.prototype, "hasBatch", {
                /**
                 * Indicates if the current query has a batch associated
                 *
                 */
                get: function () {
                    return this._batch !== null;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Queryable.prototype, "parentUrl", {
                /**
                 * Gets the parent url used when creating this instance
                 *
                 */
                get: function () {
                    return this._parentUrl;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Queryable.prototype, "query", {
                /**
                 * Provides access to the query builder for this url
                 *
                 */
                get: function () {
                    return this._query;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Creates a new instance of the supplied factory and extends this into that new instance
             *
             * @param factory constructor for the new queryable
             */
            Queryable.prototype.as = function (factory) {
                var o = new factory(this._url, null);
                return util_1.Util.extend(o, this, true);
            };
            /**
             * Adds this query to the supplied batch
             *
             * @example
             * ```
             *
             * let b = pnp.sp.createBatch();
             * pnp.sp.web.inBatch(b).get().then(...);
             * b.execute().then(...)
             * ```
             */
            Queryable.prototype.inBatch = function (batch) {
                if (this._batch !== null) {
                    throw new exceptions_1.AlreadyInBatchException();
                }
                this._batch = batch;
                return this;
            };
            /**
             * Enables caching for this request
             *
             * @param options Defines the options used when caching this request
             */
            Queryable.prototype.usingCaching = function (options) {
                if (!pnplibconfig_1.RuntimeConfig.globalCacheDisable) {
                    this._useCaching = true;
                    this._cachingOptions = options;
                }
                return this;
            };
            /**
             * Gets the currentl url, made absolute based on the availability of the _spPageContextInfo object
             *
             */
            Queryable.prototype.toUrl = function () {
                return this._url;
            };
            /**
             * Gets the full url with query information
             *
             */
            Queryable.prototype.toUrlAndQuery = function () {
                var _this = this;
                var url = this.toUrl();
                if (this._query.count() > 0) {
                    url += "?" + this._query.getKeys().map(function (key) { return key + "=" + _this._query.get(key); }).join("&");
                }
                return url;
            };
            /**
             * Gets a parent for this instance as specified
             *
             * @param factory The contructor for the class to create
             */
            Queryable.prototype.getParent = function (factory, baseUrl, path) {
                if (baseUrl === void 0) { baseUrl = this.parentUrl; }
                var parent = new factory(baseUrl, path);
                var target = this.query.get("@target");
                if (target !== null) {
                    parent.query.add("@target", target);
                }
                return parent;
            };
            /**
             * Executes the currently built request
             *
             * @param parser Allows you to specify a parser to handle the result
             * @param getOptions The options used for this request
             */
            Queryable.prototype.get = function (parser, getOptions) {
                if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
                if (getOptions === void 0) { getOptions = {}; }
                return this.toRequestContext("GET", getOptions, parser).then(function (context) { return queryablerequest_1.pipe(context); });
            };
            Queryable.prototype.getAs = function (parser, getOptions) {
                if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
                if (getOptions === void 0) { getOptions = {}; }
                return this.toRequestContext("GET", getOptions, parser).then(function (context) { return queryablerequest_1.pipe(context); });
            };
            Queryable.prototype.post = function (postOptions, parser) {
                if (postOptions === void 0) { postOptions = {}; }
                if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
                return this.toRequestContext("POST", postOptions, parser).then(function (context) { return queryablerequest_1.pipe(context); });
            };
            Queryable.prototype.postAs = function (postOptions, parser) {
                if (postOptions === void 0) { postOptions = {}; }
                if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
                return this.toRequestContext("POST", postOptions, parser).then(function (context) { return queryablerequest_1.pipe(context); });
            };
            Queryable.prototype.patch = function (patchOptions, parser) {
                if (patchOptions === void 0) { patchOptions = {}; }
                if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
                return this.toRequestContext("PATCH", patchOptions, parser).then(function (context) { return queryablerequest_1.pipe(context); });
            };
            Queryable.prototype.delete = function (deleteOptions, parser) {
                if (deleteOptions === void 0) { deleteOptions = {}; }
                if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
                return this.toRequestContext("DELETE", deleteOptions, parser).then(function (context) { return queryablerequest_1.pipe(context); });
            };
            Queryable.prototype.toRequestContext = function (verb, options, parser) {
                var _this = this;
                if (options === void 0) { options = {}; }
                var dependencyDispose = this.hasBatch ? this.addBatchDependency() : function () { return; };
                return util_1.Util.toAbsoluteUrl(this.toUrlAndQuery()).then(function (url) {
                    // build our request context
                    var context = {
                        batch: _this._batch,
                        batchDependency: dependencyDispose,
                        cachingOptions: _this._cachingOptions,
                        isBatched: _this.hasBatch,
                        isCached: _this._useCaching,
                        options: options,
                        parser: parser,
                        requestAbsoluteUrl: url,
                        requestId: util_1.Util.getGUID(),
                        verb: verb,
                    };
                    return context;
                });
            };
            return Queryable;
        }());
        exports.Queryable = Queryable;
        /**
         * Represents a REST collection which can be filtered, paged, and selected
         *
         */
        var QueryableCollection = (function (_super) {
            __extends(QueryableCollection, _super);
            function QueryableCollection() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            /**
             * Filters the returned collection (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_supported)
             *
             * @param filter The string representing the filter query
             */
            QueryableCollection.prototype.filter = function (filter) {
                this._query.add("$filter", filter);
                return this;
            };
            /**
             * Choose which fields to return
             *
             * @param selects One or more fields to return
             */
            QueryableCollection.prototype.select = function () {
                var selects = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    selects[_i] = arguments[_i];
                }
                this._query.add("$select", selects.join(","));
                return this;
            };
            /**
             * Expands fields such as lookups to get additional data
             *
             * @param expands The Fields for which to expand the values
             */
            QueryableCollection.prototype.expand = function () {
                var expands = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    expands[_i] = arguments[_i];
                }
                this._query.add("$expand", expands.join(","));
                return this;
            };
            /**
             * Orders based on the supplied fields ascending
             *
             * @param orderby The name of the field to sort on
             * @param ascending If false DESC is appended, otherwise ASC (default)
             */
            QueryableCollection.prototype.orderBy = function (orderBy, ascending) {
                if (ascending === void 0) { ascending = true; }
                var keys = this._query.getKeys();
                var query = [];
                var asc = ascending ? " asc" : " desc";
                for (var i = 0; i < keys.length; i++) {
                    if (keys[i] === "$orderby") {
                        query.push(this._query.get("$orderby"));
                        break;
                    }
                }
                query.push("" + orderBy + asc);
                this._query.add("$orderby", query.join(","));
                return this;
            };
            /**
             * Skips the specified number of items
             *
             * @param skip The number of items to skip
             */
            QueryableCollection.prototype.skip = function (skip) {
                this._query.add("$skip", skip.toString());
                return this;
            };
            /**
             * Limits the query to only return the specified number of items
             *
             * @param top The query row limit
             */
            QueryableCollection.prototype.top = function (top) {
                this._query.add("$top", top.toString());
                return this;
            };
            return QueryableCollection;
        }(Queryable));
        exports.QueryableCollection = QueryableCollection;
        /**
         * Represents an instance that can be selected
         *
         */
        var QueryableInstance = (function (_super) {
            __extends(QueryableInstance, _super);
            function QueryableInstance() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            /**
             * Choose which fields to return
             *
             * @param selects One or more fields to return
             */
            QueryableInstance.prototype.select = function () {
                var selects = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    selects[_i] = arguments[_i];
                }
                this._query.add("$select", selects.join(","));
                return this;
            };
            /**
             * Expands fields such as lookups to get additional data
             *
             * @param expands The Fields for which to expand the values
             */
            QueryableInstance.prototype.expand = function () {
                var expands = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    expands[_i] = arguments[_i];
                }
                this._query.add("$expand", expands.join(","));
                return this;
            };
            return QueryableInstance;
        }(Queryable));
        exports.QueryableInstance = QueryableInstance;


        /***/
    }),
    /* 2 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var logging_1 = __webpack_require__(5);
        function defaultLog(error) {
            logging_1.Logger.log({ data: {}, level: logging_1.LogLevel.Error, message: "[" + error.name + "]::" + error.message });
        }
        /**
         * Represents an exception with an HttpClient request
         *
         */
        var ProcessHttpClientResponseException = (function (_super) {
            __extends(ProcessHttpClientResponseException, _super);
            function ProcessHttpClientResponseException(status, statusText, data) {
                var _this = _super.call(this, "Error making HttpClient request in queryable: [" + status + "] " + statusText) || this;
                _this.status = status;
                _this.statusText = statusText;
                _this.data = data;
                _this.name = "ProcessHttpClientResponseException";
                logging_1.Logger.log({ data: _this.data, level: logging_1.LogLevel.Error, message: _this.message });
                return _this;
            }
            return ProcessHttpClientResponseException;
        }(Error));
        exports.ProcessHttpClientResponseException = ProcessHttpClientResponseException;
        var NoCacheAvailableException = (function (_super) {
            __extends(NoCacheAvailableException, _super);
            function NoCacheAvailableException(msg) {
                if (msg === void 0) { msg = "Cannot create a caching configuration provider since cache is not available."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "NoCacheAvailableException";
                defaultLog(_this);
                return _this;
            }
            return NoCacheAvailableException;
        }(Error));
        exports.NoCacheAvailableException = NoCacheAvailableException;
        var APIUrlException = (function (_super) {
            __extends(APIUrlException, _super);
            function APIUrlException(msg) {
                if (msg === void 0) { msg = "Unable to determine API url."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "APIUrlException";
                defaultLog(_this);
                return _this;
            }
            return APIUrlException;
        }(Error));
        exports.APIUrlException = APIUrlException;
        var AuthUrlException = (function (_super) {
            __extends(AuthUrlException, _super);
            function AuthUrlException(data, msg) {
                if (msg === void 0) { msg = "Auth URL Endpoint could not be determined from data. Data logged."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "APIUrlException";
                logging_1.Logger.log({ data: data, level: logging_1.LogLevel.Error, message: _this.message });
                return _this;
            }
            return AuthUrlException;
        }(Error));
        exports.AuthUrlException = AuthUrlException;
        var NodeFetchClientUnsupportedException = (function (_super) {
            __extends(NodeFetchClientUnsupportedException, _super);
            function NodeFetchClientUnsupportedException(msg) {
                if (msg === void 0) { msg = "Using NodeFetchClient in the browser is not supported."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "NodeFetchClientUnsupportedException";
                defaultLog(_this);
                return _this;
            }
            return NodeFetchClientUnsupportedException;
        }(Error));
        exports.NodeFetchClientUnsupportedException = NodeFetchClientUnsupportedException;
        var SPRequestExecutorUndefinedException = (function (_super) {
            __extends(SPRequestExecutorUndefinedException, _super);
            function SPRequestExecutorUndefinedException() {
                var _this = this;
                var msg = [
                    "SP.RequestExecutor is undefined. ",
                    "Load the SP.RequestExecutor.js library (/_layouts/15/SP.RequestExecutor.js) before loading the PnP JS Core library.",
                ].join(" ");
                _this = _super.call(this, msg) || this;
                _this.name = "SPRequestExecutorUndefinedException";
                defaultLog(_this);
                return _this;
            }
            return SPRequestExecutorUndefinedException;
        }(Error));
        exports.SPRequestExecutorUndefinedException = SPRequestExecutorUndefinedException;
        var MaxCommentLengthException = (function (_super) {
            __extends(MaxCommentLengthException, _super);
            function MaxCommentLengthException(msg) {
                if (msg === void 0) { msg = "The maximum comment length is 1023 characters."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "MaxCommentLengthException";
                defaultLog(_this);
                return _this;
            }
            return MaxCommentLengthException;
        }(Error));
        exports.MaxCommentLengthException = MaxCommentLengthException;
        var NotSupportedInBatchException = (function (_super) {
            __extends(NotSupportedInBatchException, _super);
            function NotSupportedInBatchException(operation) {
                if (operation === void 0) { operation = "This operation"; }
                var _this = _super.call(this, operation + " is not supported as part of a batch.") || this;
                _this.name = "NotSupportedInBatchException";
                defaultLog(_this);
                return _this;
            }
            return NotSupportedInBatchException;
        }(Error));
        exports.NotSupportedInBatchException = NotSupportedInBatchException;
        var ODataIdException = (function (_super) {
            __extends(ODataIdException, _super);
            function ODataIdException(data, msg) {
                if (msg === void 0) { msg = "Could not extract odata id in object, you may be using nometadata. Object data logged to logger."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "ODataIdException";
                logging_1.Logger.log({ data: data, level: logging_1.LogLevel.Error, message: _this.message });
                return _this;
            }
            return ODataIdException;
        }(Error));
        exports.ODataIdException = ODataIdException;
        var BatchParseException = (function (_super) {
            __extends(BatchParseException, _super);
            function BatchParseException(msg) {
                var _this = _super.call(this, msg) || this;
                _this.name = "BatchParseException";
                defaultLog(_this);
                return _this;
            }
            return BatchParseException;
        }(Error));
        exports.BatchParseException = BatchParseException;
        var AlreadyInBatchException = (function (_super) {
            __extends(AlreadyInBatchException, _super);
            function AlreadyInBatchException(msg) {
                if (msg === void 0) { msg = "This query is already part of a batch."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "AlreadyInBatchException";
                defaultLog(_this);
                return _this;
            }
            return AlreadyInBatchException;
        }(Error));
        exports.AlreadyInBatchException = AlreadyInBatchException;
        var FunctionExpectedException = (function (_super) {
            __extends(FunctionExpectedException, _super);
            function FunctionExpectedException(msg) {
                if (msg === void 0) { msg = "This query is already part of a batch."; }
                var _this = _super.call(this, msg) || this;
                _this.name = "FunctionExpectedException";
                defaultLog(_this);
                return _this;
            }
            return FunctionExpectedException;
        }(Error));
        exports.FunctionExpectedException = FunctionExpectedException;
        var UrlException = (function (_super) {
            __extends(UrlException, _super);
            function UrlException(msg) {
                var _this = _super.call(this, msg) || this;
                _this.name = "UrlException";
                defaultLog(_this);
                return _this;
            }
            return UrlException;
        }(Error));
        exports.UrlException = UrlException;


        /***/
    }),
    /* 3 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        var logging_1 = __webpack_require__(5);
        var httpclient_1 = __webpack_require__(10);
        var pnplibconfig_1 = __webpack_require__(4);
        var exceptions_1 = __webpack_require__(2);
        var exceptions_2 = __webpack_require__(2);
        function extractOdataId(candidate) {
            if (candidate.hasOwnProperty("odata.id")) {
                return candidate["odata.id"];
            }
            else if (candidate.hasOwnProperty("__metadata") && candidate.__metadata.hasOwnProperty("id")) {
                return candidate.__metadata.id;
            }
            else {
                throw new exceptions_1.ODataIdException(candidate);
            }
        }
        exports.extractOdataId = extractOdataId;
        var ODataParserBase = (function () {
            function ODataParserBase() {
            }
            ODataParserBase.prototype.parse = function (r) {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    if (_this.handleError(r, reject)) {
                        if ((r.headers.has("Content-Length") && parseFloat(r.headers.get("Content-Length")) === 0) || r.status === 204) {
                            resolve({});
                        }
                        else {
                            r.json().then(function (json) { return resolve(_this.parseODataJSON(json)); });
                        }
                    }
                });
            };
            ODataParserBase.prototype.handleError = function (r, reject) {
                if (!r.ok) {
                    r.json().then(function (json) {
                        // include the headers as they contain diagnostic information
                        var data = {
                            responseBody: json,
                            responseHeaders: r.headers,
                        };
                        reject(new exceptions_2.ProcessHttpClientResponseException(r.status, r.statusText, data));
                    });
                }
                return r.ok;
            };
            ODataParserBase.prototype.parseODataJSON = function (json) {
                var result = json;
                if (json.hasOwnProperty("d")) {
                    if (json.d.hasOwnProperty("results")) {
                        result = json.d.results;
                    }
                    else {
                        result = json.d;
                    }
                }
                else if (json.hasOwnProperty("value")) {
                    result = json.value;
                }
                return result;
            };
            return ODataParserBase;
        }());
        exports.ODataParserBase = ODataParserBase;
        var ODataDefaultParser = (function (_super) {
            __extends(ODataDefaultParser, _super);
            function ODataDefaultParser() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            return ODataDefaultParser;
        }(ODataParserBase));
        exports.ODataDefaultParser = ODataDefaultParser;
        var ODataRawParserImpl = (function () {
            function ODataRawParserImpl() {
            }
            ODataRawParserImpl.prototype.parse = function (r) {
                return r.json();
            };
            return ODataRawParserImpl;
        }());
        exports.ODataRawParserImpl = ODataRawParserImpl;
        var ODataValueParserImpl = (function (_super) {
            __extends(ODataValueParserImpl, _super);
            function ODataValueParserImpl() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            ODataValueParserImpl.prototype.parse = function (r) {
                return _super.prototype.parse.call(this, r).then(function (d) { return d; });
            };
            return ODataValueParserImpl;
        }(ODataParserBase));
        var ODataEntityParserImpl = (function (_super) {
            __extends(ODataEntityParserImpl, _super);
            function ODataEntityParserImpl(factory) {
                var _this = _super.call(this) || this;
                _this.factory = factory;
                return _this;
            }
            ODataEntityParserImpl.prototype.parse = function (r) {
                var _this = this;
                return _super.prototype.parse.call(this, r).then(function (d) {
                    var o = new _this.factory(getEntityUrl(d), null);
                    return util_1.Util.extend(o, d);
                });
            };
            return ODataEntityParserImpl;
        }(ODataParserBase));
        var ODataEntityArrayParserImpl = (function (_super) {
            __extends(ODataEntityArrayParserImpl, _super);
            function ODataEntityArrayParserImpl(factory) {
                var _this = _super.call(this) || this;
                _this.factory = factory;
                return _this;
            }
            ODataEntityArrayParserImpl.prototype.parse = function (r) {
                var _this = this;
                return _super.prototype.parse.call(this, r).then(function (d) {
                    return d.map(function (v) {
                        var o = new _this.factory(getEntityUrl(v), null);
                        return util_1.Util.extend(o, v);
                    });
                });
            };
            return ODataEntityArrayParserImpl;
        }(ODataParserBase));
        function getEntityUrl(entity) {
            if (entity.hasOwnProperty("odata.editLink")) {
                // we are dealign with minimal metadata (default)
                return util_1.Util.combinePaths("_api", entity["odata.editLink"]);
            }
            else if (entity.hasOwnProperty("__metadata")) {
                // we are dealing with verbose, which has an absolute uri
                return entity.__metadata.uri;
            }
            else {
                // we are likely dealing with nometadata, so don't error but we won't be able to
                // chain off these objects
                logging_1.Logger.write("No uri information found in ODataEntity parsing, chaining will fail for this object.", logging_1.LogLevel.Warning);
                return "";
            }
        }
        exports.ODataRaw = new ODataRawParserImpl();
        function ODataValue() {
            return new ODataValueParserImpl();
        }
        exports.ODataValue = ODataValue;
        function ODataEntity(factory) {
            return new ODataEntityParserImpl(factory);
        }
        exports.ODataEntity = ODataEntity;
        function ODataEntityArray(factory) {
            return new ODataEntityArrayParserImpl(factory);
        }
        exports.ODataEntityArray = ODataEntityArray;
        /**
         * Manages a batch of OData operations
         */
        var ODataBatch = (function () {
            function ODataBatch(baseUrl, _batchId) {
                if (_batchId === void 0) { _batchId = util_1.Util.getGUID(); }
                this.baseUrl = baseUrl;
                this._batchId = _batchId;
                this._requests = [];
                this._dependencies = Promise.resolve();
            }
            /**
             * Adds a request to a batch (not designed for public use)
             *
             * @param url The full url of the request
             * @param method The http method GET, POST, etc
             * @param options Any options to include in the request
             * @param parser The parser that will hadle the results of the request
             */
            ODataBatch.prototype.add = function (url, method, options, parser) {
                var info = {
                    method: method.toUpperCase(),
                    options: options,
                    parser: parser,
                    reject: null,
                    resolve: null,
                    url: url,
                };
                var p = new Promise(function (resolve, reject) {
                    info.resolve = resolve;
                    info.reject = reject;
                });
                this._requests.push(info);
                return p;
            };
            /**
             * Adds a dependency insuring that some set of actions will occur before a batch is processed.
             * MUST be cleared using the returned resolve delegate to allow batches to run
             */
            ODataBatch.prototype.addDependency = function () {
                var resolver;
                var promise = new Promise(function (resolve) {
                    resolver = resolve;
                });
                this._dependencies = this._dependencies.then(function () { return promise; });
                return resolver;
            };
            /**
             * Execute the current batch and resolve the associated promises
             *
             * @returns A promise which will be resolved once all of the batch's child promises have resolved
             */
            ODataBatch.prototype.execute = function () {
                var _this = this;
                return this._dependencies.then(function () { return _this.executeImpl(); });
            };
            ODataBatch.prototype.executeImpl = function () {
                var _this = this;
                logging_1.Logger.write("Executing batch with " + this._requests.length + " requests.", logging_1.LogLevel.Info);
                // if we don't have any requests, don't bother sending anything
                // this could be due to caching further upstream, or just an empty batch
                if (this._requests.length < 1) {
                    logging_1.Logger.write("Resolving empty batch.", logging_1.LogLevel.Info);
                    return Promise.resolve();
                }
                // creating the client here allows the url to be populated for nodejs client as well as potentially
                // any other hacks needed for other types of clients. Essentially allows the absoluteRequestUrl
                // below to be correct
                var client = new httpclient_1.HttpClient();
                // due to timing we need to get the absolute url here so we can use it for all the individual requests
                // and for sending the entire batch
                return util_1.Util.toAbsoluteUrl(this.baseUrl).then(function (absoluteRequestUrl) {
                    // build all the requests, send them, pipe results in order to parsers
                    var batchBody = [];
                    var currentChangeSetId = "";
                    _this._requests.map(function (reqInfo) {
                        if (reqInfo.method === "GET") {
                            if (currentChangeSetId.length > 0) {
                                // end an existing change set
                                batchBody.push("--changeset_" + currentChangeSetId + "--\n\n");
                                currentChangeSetId = "";
                            }
                            batchBody.push("--batch_" + _this._batchId + "\n");
                        }
                        else {
                            if (currentChangeSetId.length < 1) {
                                // start new change set
                                currentChangeSetId = util_1.Util.getGUID();
                                batchBody.push("--batch_" + _this._batchId + "\n");
                                batchBody.push("Content-Type: multipart/mixed; boundary=\"changeset_" + currentChangeSetId + "\"\n\n");
                            }
                            batchBody.push("--changeset_" + currentChangeSetId + "\n");
                        }
                        // common batch part prefix
                        batchBody.push("Content-Type: application/http\n");
                        batchBody.push("Content-Transfer-Encoding: binary\n\n");
                        var headers = {
                            "Accept": "application/json;",
                        };
                        // this is the url of the individual request within the batch
                        var url = util_1.Util.isUrlAbsolute(reqInfo.url) ? reqInfo.url : util_1.Util.combinePaths(absoluteRequestUrl, reqInfo.url);
                        logging_1.Logger.write("Adding request " + reqInfo.method + " " + url + " to batch.", logging_1.LogLevel.Verbose);
                        if (reqInfo.method !== "GET") {
                            var method = reqInfo.method;
                            if (reqInfo.hasOwnProperty("options") && reqInfo.options.hasOwnProperty("headers") && typeof reqInfo.options.headers["X-HTTP-Method"] !== "undefined") {
                                method = reqInfo.options.headers["X-HTTP-Method"];
                                delete reqInfo.options.headers["X-HTTP-Method"];
                            }
                            batchBody.push(method + " " + url + " HTTP/1.1\n");
                            headers = util_1.Util.extend(headers, { "Content-Type": "application/json;odata=verbose;charset=utf-8" });
                        }
                        else {
                            batchBody.push(reqInfo.method + " " + url + " HTTP/1.1\n");
                        }
                        if (typeof pnplibconfig_1.RuntimeConfig.headers !== "undefined") {
                            headers = util_1.Util.extend(headers, pnplibconfig_1.RuntimeConfig.headers);
                        }
                        if (reqInfo.options && reqInfo.options.headers) {
                            headers = util_1.Util.extend(headers, reqInfo.options.headers);
                        }
                        for (var name_1 in headers) {
                            if (headers.hasOwnProperty(name_1)) {
                                batchBody.push(name_1 + ": " + headers[name_1] + "\n");
                            }
                        }
                        batchBody.push("\n");
                        if (reqInfo.options.body) {
                            batchBody.push(reqInfo.options.body + "\n\n");
                        }
                    });
                    if (currentChangeSetId.length > 0) {
                        // Close the changeset
                        batchBody.push("--changeset_" + currentChangeSetId + "--\n\n");
                        currentChangeSetId = "";
                    }
                    batchBody.push("--batch_" + _this._batchId + "--\n");
                    var batchHeaders = {
                        "Content-Type": "multipart/mixed; boundary=batch_" + _this._batchId,
                    };
                    var batchOptions = {
                        "body": batchBody.join(""),
                        "headers": batchHeaders,
                    };
                    logging_1.Logger.write("Sending batch request.", logging_1.LogLevel.Info);
                    return client.post(util_1.Util.combinePaths(absoluteRequestUrl, "/_api/$batch"), batchOptions)
                        .then(function (r) { return r.text(); })
                        .then(_this._parseResponse)
                        .then(function (responses) {
                            if (responses.length !== _this._requests.length) {
                                throw new exceptions_1.BatchParseException("Could not properly parse responses to match requests in batch.");
                            }
                            logging_1.Logger.write("Resolving batched requests.", logging_1.LogLevel.Info);
                            return responses.reduce(function (chain, response, index) {
                                var request = _this._requests[index];
                                logging_1.Logger.write("Resolving request " + request.method + " " + request.url + ".", logging_1.LogLevel.Verbose);
                                return chain.then(function (_) { return request.parser.parse(response).then(request.resolve).catch(request.reject); });
                            }, Promise.resolve());
                        });
                });
            };
            /**
             * Parses the response from a batch request into an array of Response instances
             *
             * @param body Text body of the response from the batch request
             */
            ODataBatch.prototype._parseResponse = function (body) {
                return new Promise(function (resolve, reject) {
                    var responses = [];
                    var header = "--batchresponse_";
                    // Ex. "HTTP/1.1 500 Internal Server Error"
                    var statusRegExp = new RegExp("^HTTP/[0-9.]+ +([0-9]+) +(.*)", "i");
                    var lines = body.split("\n");
                    var state = "batch";
                    var status;
                    var statusText;
                    for (var i = 0; i < lines.length; ++i) {
                        var line = lines[i];
                        switch (state) {
                            case "batch":
                                if (line.substr(0, header.length) === header) {
                                    state = "batchHeaders";
                                }
                                else {
                                    if (line.trim() !== "") {
                                        throw new exceptions_1.BatchParseException("Invalid response, line " + i);
                                    }
                                }
                                break;
                            case "batchHeaders":
                                if (line.trim() === "") {
                                    state = "status";
                                }
                                break;
                            case "status":
                                var parts = statusRegExp.exec(line);
                                if (parts.length !== 3) {
                                    throw new exceptions_1.BatchParseException("Invalid status, line " + i);
                                }
                                status = parseInt(parts[1], 10);
                                statusText = parts[2];
                                state = "statusHeaders";
                                break;
                            case "statusHeaders":
                                if (line.trim() === "") {
                                    state = "body";
                                }
                                break;
                            case "body":
                                responses.push((status === 204) ? new Response() : new Response(line, { status: status, statusText: statusText }));
                                state = "batch";
                                break;
                        }
                    }
                    if (state !== "status") {
                        reject(new exceptions_1.BatchParseException("Unexpected end of input"));
                    }
                    resolve(responses);
                });
            };
            return ODataBatch;
        }());
        exports.ODataBatch = ODataBatch;
        var TextFileParser = (function () {
            function TextFileParser() {
            }
            TextFileParser.prototype.parse = function (r) {
                return r.text();
            };
            return TextFileParser;
        }());
        exports.TextFileParser = TextFileParser;
        var BlobFileParser = (function () {
            function BlobFileParser() {
            }
            BlobFileParser.prototype.parse = function (r) {
                return r.blob();
            };
            return BlobFileParser;
        }());
        exports.BlobFileParser = BlobFileParser;
        var JSONFileParser = (function () {
            function JSONFileParser() {
            }
            JSONFileParser.prototype.parse = function (r) {
                return r.json();
            };
            return JSONFileParser;
        }());
        exports.JSONFileParser = JSONFileParser;
        var BufferFileParser = (function () {
            function BufferFileParser() {
            }
            BufferFileParser.prototype.parse = function (r) {
                if (util_1.Util.isFunction(r.arrayBuffer)) {
                    return r.arrayBuffer();
                }
                return r.buffer();
            };
            return BufferFileParser;
        }());
        exports.BufferFileParser = BufferFileParser;


        /***/
    }),
    /* 4 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var fetchclient_1 = __webpack_require__(18);
        var RuntimeConfigImpl = (function () {
            function RuntimeConfigImpl() {
                // these are our default values for the library
                this._headers = null;
                this._defaultCachingStore = "session";
                this._defaultCachingTimeoutSeconds = 30;
                this._globalCacheDisable = false;
                this._fetchClientFactory = function () { return new fetchclient_1.FetchClient(); };
                this._baseUrl = null;
                this._spfxContext = null;
            }
            RuntimeConfigImpl.prototype.set = function (config) {
                if (config.hasOwnProperty("headers")) {
                    this._headers = config.headers;
                }
                if (config.hasOwnProperty("globalCacheDisable")) {
                    this._globalCacheDisable = config.globalCacheDisable;
                }
                if (config.hasOwnProperty("defaultCachingStore")) {
                    this._defaultCachingStore = config.defaultCachingStore;
                }
                if (config.hasOwnProperty("defaultCachingTimeoutSeconds")) {
                    this._defaultCachingTimeoutSeconds = config.defaultCachingTimeoutSeconds;
                }
                if (config.hasOwnProperty("fetchClientFactory")) {
                    this._fetchClientFactory = config.fetchClientFactory;
                }
                if (config.hasOwnProperty("baseUrl")) {
                    this._baseUrl = config.baseUrl;
                }
                if (config.hasOwnProperty("spFXContext")) {
                    this._spfxContext = config.spfxContext;
                }
            };
            Object.defineProperty(RuntimeConfigImpl.prototype, "headers", {
                get: function () {
                    return this._headers;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RuntimeConfigImpl.prototype, "defaultCachingStore", {
                get: function () {
                    return this._defaultCachingStore;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RuntimeConfigImpl.prototype, "defaultCachingTimeoutSeconds", {
                get: function () {
                    return this._defaultCachingTimeoutSeconds;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RuntimeConfigImpl.prototype, "globalCacheDisable", {
                get: function () {
                    return this._globalCacheDisable;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RuntimeConfigImpl.prototype, "fetchClientFactory", {
                get: function () {
                    return this._fetchClientFactory;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RuntimeConfigImpl.prototype, "baseUrl", {
                get: function () {
                    if (this._baseUrl !== null) {
                        return this._baseUrl;
                    }
                    else if (this._spfxContext !== null) {
                        return this._spfxContext.pageContext.web.absoluteUrl;
                    }
                    return null;
                },
                enumerable: true,
                configurable: true
            });
            return RuntimeConfigImpl;
        }());
        exports.RuntimeConfigImpl = RuntimeConfigImpl;
        var _runtimeConfig = new RuntimeConfigImpl();
        exports.RuntimeConfig = _runtimeConfig;
        function setRuntimeConfig(config) {
            _runtimeConfig.set(config);
        }
        exports.setRuntimeConfig = setRuntimeConfig;


        /***/
    }),
    /* 5 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        /**
         * A set of logging levels
         *
         */
        var LogLevel;
        (function (LogLevel) {
            LogLevel[LogLevel["Verbose"] = 0] = "Verbose";
            LogLevel[LogLevel["Info"] = 1] = "Info";
            LogLevel[LogLevel["Warning"] = 2] = "Warning";
            LogLevel[LogLevel["Error"] = 3] = "Error";
            LogLevel[LogLevel["Off"] = 99] = "Off";
        })(LogLevel = exports.LogLevel || (exports.LogLevel = {}));
        /**
         * Class used to subscribe ILogListener and log messages throughout an application
         *
         */
        var Logger = (function () {
            function Logger() {
            }
            Object.defineProperty(Logger, "activeLogLevel", {
                get: function () {
                    return Logger.instance.activeLogLevel;
                },
                set: function (value) {
                    Logger.instance.activeLogLevel = value;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Logger, "instance", {
                get: function () {
                    if (typeof Logger._instance === "undefined" || Logger._instance === null) {
                        Logger._instance = new LoggerImpl();
                    }
                    return Logger._instance;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Adds ILogListener instances to the set of subscribed listeners
             *
             * @param listeners One or more listeners to subscribe to this log
             */
            Logger.subscribe = function () {
                var listeners = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    listeners[_i] = arguments[_i];
                }
                listeners.map(function (listener) { return Logger.instance.subscribe(listener); });
            };
            /**
             * Clears the subscribers collection, returning the collection before modifiction
             */
            Logger.clearSubscribers = function () {
                return Logger.instance.clearSubscribers();
            };
            Object.defineProperty(Logger, "count", {
                /**
                 * Gets the current subscriber count
                 */
                get: function () {
                    return Logger.instance.count;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Writes the supplied string to the subscribed listeners
             *
             * @param message The message to write
             * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Verbose)
             */
            Logger.write = function (message, level) {
                if (level === void 0) { level = LogLevel.Verbose; }
                Logger.instance.log({ level: level, message: message });
            };
            /**
             * Writes the supplied string to the subscribed listeners
             *
             * @param json The json object to stringify and write
             * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Verbose)
             */
            Logger.writeJSON = function (json, level) {
                if (level === void 0) { level = LogLevel.Verbose; }
                Logger.instance.log({ level: level, message: JSON.stringify(json) });
            };
            /**
             * Logs the supplied entry to the subscribed listeners
             *
             * @param entry The message to log
             */
            Logger.log = function (entry) {
                Logger.instance.log(entry);
            };
            /**
             * Logs performance tracking data for the the execution duration of the supplied function using console.profile
             *
             * @param name The name of this profile boundary
             * @param f The function to execute and track within this performance boundary
             */
            Logger.measure = function (name, f) {
                return Logger.instance.measure(name, f);
            };
            return Logger;
        }());
        exports.Logger = Logger;
        var LoggerImpl = (function () {
            function LoggerImpl(activeLogLevel, subscribers) {
                if (activeLogLevel === void 0) { activeLogLevel = LogLevel.Warning; }
                if (subscribers === void 0) { subscribers = []; }
                this.activeLogLevel = activeLogLevel;
                this.subscribers = subscribers;
            }
            LoggerImpl.prototype.subscribe = function (listener) {
                this.subscribers.push(listener);
            };
            LoggerImpl.prototype.clearSubscribers = function () {
                var s = this.subscribers.slice(0);
                this.subscribers.length = 0;
                return s;
            };
            Object.defineProperty(LoggerImpl.prototype, "count", {
                get: function () {
                    return this.subscribers.length;
                },
                enumerable: true,
                configurable: true
            });
            LoggerImpl.prototype.write = function (message, level) {
                if (level === void 0) { level = LogLevel.Verbose; }
                this.log({ level: level, message: message });
            };
            LoggerImpl.prototype.log = function (entry) {
                if (typeof entry === "undefined" || entry.level < this.activeLogLevel) {
                    return;
                }
                this.subscribers.map(function (subscriber) { return subscriber.log(entry); });
            };
            LoggerImpl.prototype.measure = function (name, f) {
                console.profile(name);
                try {
                    return f();
                }
                finally {
                    console.profileEnd();
                }
            };
            return LoggerImpl;
        }());
        /**
         * Implementation of ILogListener which logs to the browser console
         *
         */
        var ConsoleListener = (function () {
            function ConsoleListener() {
            }
            /**
             * Any associated data that a given logging listener may choose to log or ignore
             *
             * @param entry The information to be logged
             */
            ConsoleListener.prototype.log = function (entry) {
                var msg = this.format(entry);
                switch (entry.level) {
                    case LogLevel.Verbose:
                    case LogLevel.Info:
                        console.log(msg);
                        break;
                    case LogLevel.Warning:
                        console.warn(msg);
                        break;
                    case LogLevel.Error:
                        console.error(msg);
                        break;
                }
            };
            /**
             * Formats the message
             *
             * @param entry The information to format into a string
             */
            ConsoleListener.prototype.format = function (entry) {
                return "Message: " + entry.message + " Data: " + JSON.stringify(entry.data);
            };
            return ConsoleListener;
        }());
        exports.ConsoleListener = ConsoleListener;
        /**
         * Implementation of ILogListener which logs to the supplied function
         *
         */
        var FunctionListener = (function () {
            /**
             * Creates a new instance of the FunctionListener class
             *
             * @constructor
             * @param  method The method to which any logging data will be passed
             */
            function FunctionListener(method) {
                this.method = method;
            }
            /**
             * Any associated data that a given logging listener may choose to log or ignore
             *
             * @param entry The information to be logged
             */
            FunctionListener.prototype.log = function (entry) {
                this.method(entry);
            };
            return FunctionListener;
        }());
        exports.FunctionListener = FunctionListener;


        /***/
    }),
    /* 6 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        /**
         * Generic dictionary
         */
        var Dictionary = (function () {
            /**
             * Creates a new instance of the Dictionary<T> class
             *
             * @constructor
             */
            function Dictionary(keys, values) {
                if (keys === void 0) { keys = []; }
                if (values === void 0) { values = []; }
                this.keys = keys;
                this.values = values;
            }
            /**
             * Gets a value from the collection using the specified key
             *
             * @param key The key whose value we want to return, returns null if the key does not exist
             */
            Dictionary.prototype.get = function (key) {
                var index = this.keys.indexOf(key);
                if (index < 0) {
                    return null;
                }
                return this.values[index];
            };
            /**
             * Adds the supplied key and value to the dictionary
             *
             * @param key The key to add
             * @param o The value to add
             */
            Dictionary.prototype.add = function (key, o) {
                var index = this.keys.indexOf(key);
                if (index > -1) {
                    this.values[index] = o;
                }
                else {
                    this.keys.push(key);
                    this.values.push(o);
                }
            };
            /**
             * Merges the supplied typed hash into this dictionary instance. Existing values are updated and new ones are created as appropriate.
             */
            Dictionary.prototype.merge = function (source) {
                var _this = this;
                if ("getKeys" in source) {
                    var sourceAsDictionary_1 = source;
                    sourceAsDictionary_1.getKeys().map(function (key) {
                        _this.add(key, sourceAsDictionary_1.get(key));
                    });
                }
                else {
                    var sourceAsHash = source;
                    for (var key in sourceAsHash) {
                        if (sourceAsHash.hasOwnProperty(key)) {
                            this.add(key, sourceAsHash[key]);
                        }
                    }
                }
            };
            /**
             * Removes a value from the dictionary
             *
             * @param key The key of the key/value pair to remove. Returns null if the key was not found.
             */
            Dictionary.prototype.remove = function (key) {
                var index = this.keys.indexOf(key);
                if (index < 0) {
                    return null;
                }
                var val = this.values[index];
                this.keys.splice(index, 1);
                this.values.splice(index, 1);
                return val;
            };
            /**
             * Returns all the keys currently in the dictionary as an array
             */
            Dictionary.prototype.getKeys = function () {
                return this.keys;
            };
            /**
             * Returns all the values currently in the dictionary as an array
             */
            Dictionary.prototype.getValues = function () {
                return this.values;
            };
            /**
             * Clears the current dictionary
             */
            Dictionary.prototype.clear = function () {
                this.keys = [];
                this.values = [];
            };
            /**
             * Gets a count of the items currently in the dictionary
             */
            Dictionary.prototype.count = function () {
                return this.keys.length;
            };
            return Dictionary;
        }());
        exports.Dictionary = Dictionary;


        /***/
    }),
    /* 7 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var odata_1 = __webpack_require__(3);
        var util_1 = __webpack_require__(0);
        var exceptions_1 = __webpack_require__(2);
        var webparts_1 = __webpack_require__(46);
        /**
         * Describes a collection of File objects
         *
         */
        var Files = (function (_super) {
            __extends(Files, _super);
            /**
             * Creates a new instance of the Files class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Files(baseUrl, path) {
                if (path === void 0) { path = "files"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a File by filename
             *
             * @param name The name of the file, including extension.
             */
            Files.prototype.getByName = function (name) {
                var f = new File(this);
                f.concat("('" + name + "')");
                return f;
            };
            /**
             * Uploads a file.
             *
             * @param url The folder-relative url of the file.
             * @param content The file contents blob.
             * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
             * @returns The new File and the raw response.
             */
            Files.prototype.add = function (url, content, shouldOverWrite) {
                var _this = this;
                if (shouldOverWrite === void 0) { shouldOverWrite = true; }
                return new Files(this, "add(overwrite=" + shouldOverWrite + ",url='" + url + "')")
                    .post({
                        body: content,
                    }).then(function (response) {
                        return {
                            data: response,
                            file: _this.getByName(url),
                        };
                    });
            };
            /**
             * Uploads a file.
             *
             * @param url The folder-relative url of the file.
             * @param content The Blob file content to add
             * @param progress A callback function which can be used to track the progress of the upload
             * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
             * @param chunkSize The size of each file slice, in bytes (default: 10485760)
             * @returns The new File and the raw response.
             */
            Files.prototype.addChunked = function (url, content, progress, shouldOverWrite, chunkSize) {
                var _this = this;
                if (shouldOverWrite === void 0) { shouldOverWrite = true; }
                if (chunkSize === void 0) { chunkSize = 10485760; }
                var adder = new Files(this, "add(overwrite=" + shouldOverWrite + ",url='" + url + "')");
                return adder.post().then(function () { return _this.getByName(url); }).then(function (file) { return file.setContentChunked(content, progress, chunkSize); }).then(function (response) {
                    return {
                        data: response,
                        file: _this.getByName(url),
                    };
                });
            };
            /**
             * Adds a ghosted file to an existing list or document library.
             *
             * @param fileUrl The server-relative url where you want to save the file.
             * @param templateFileType The type of use to create the file.
             * @returns The template file that was added and the raw response.
             */
            Files.prototype.addTemplateFile = function (fileUrl, templateFileType) {
                var _this = this;
                return new Files(this, "addTemplateFile(urloffile='" + fileUrl + "',templatefiletype=" + templateFileType + ")")
                    .post().then(function (response) {
                        return {
                            data: response,
                            file: _this.getByName(fileUrl),
                        };
                    });
            };
            return Files;
        }(queryable_1.QueryableCollection));
        exports.Files = Files;
        /**
         * Describes a single File instance
         *
         */
        var File = (function (_super) {
            __extends(File, _super);
            /**
             * Creates a new instance of the File class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function File(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(File.prototype, "listItemAllFields", {
                /**
                 * Gets a value that specifies the list item field values for the list item corresponding to the file.
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "listItemAllFields");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(File.prototype, "versions", {
                /**
                 * Gets a collection of versions
                 *
                 */
                get: function () {
                    return new Versions(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Approves the file submitted for content approval with the specified comment.
             * Only documents in lists that are enabled for content approval can be approved.
             *
             * @param comment The comment for the approval.
             */
            File.prototype.approve = function (comment) {
                return new File(this, "approve(comment='" + comment + "')").post();
            };
            /**
             * Stops the chunk upload session without saving the uploaded data.
             * If the file doesn�t already exist in the library, the partially uploaded file will be deleted.
             * Use this in response to user action (as in a request to cancel an upload) or an error or exception.
             * Use the uploadId value that was passed to the StartUpload method that started the upload session.
             * This method is currently available only on Office 365.
             *
             * @param uploadId The unique identifier of the upload session.
             */
            File.prototype.cancelUpload = function (uploadId) {
                return new File(this, "cancelUpload(uploadId=guid'" + uploadId + "')").post();
            };
            /**
             * Checks the file in to a document library based on the check-in type.
             *
             * @param comment A comment for the check-in. Its length must be <= 1023.
             * @param checkinType The check-in type for the file.
             */
            File.prototype.checkin = function (comment, checkinType) {
                if (comment === void 0) { comment = ""; }
                if (checkinType === void 0) { checkinType = CheckinType.Major; }
                // TODO: Enforce comment length <= 1023
                return new File(this, "checkin(comment='" + comment + "',checkintype=" + checkinType + ")").post();
            };
            /**
             * Checks out the file from a document library.
             */
            File.prototype.checkout = function () {
                return new File(this, "checkout").post();
            };
            /**
             * Copies the file to the destination url.
             *
             * @param url The absolute url or server relative url of the destination file path to copy to.
             * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
             */
            File.prototype.copyTo = function (url, shouldOverWrite) {
                if (shouldOverWrite === void 0) { shouldOverWrite = true; }
                return new File(this, "copyTo(strnewurl='" + url + "',boverwrite=" + shouldOverWrite + ")").post();
            };
            /**
             * Delete this file.
             *
             * @param eTag Value used in the IF-Match header, by default "*"
             */
            File.prototype.delete = function (eTag) {
                if (eTag === void 0) { eTag = "*"; }
                return new File(this).post({
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            /**
             * Denies approval for a file that was submitted for content approval.
             * Only documents in lists that are enabled for content approval can be denied.
             *
             * @param comment The comment for the denial.
             */
            File.prototype.deny = function (comment) {
                if (comment === void 0) { comment = ""; }
                return new File(this, "deny(comment='" + comment + "')").post();
            };
            /**
             * Specifies the control set used to access, modify, or add Web Parts associated with this Web Part Page and view.
             * An exception is thrown if the file is not an ASPX page.
             *
             * @param scope The WebPartsPersonalizationScope view on the Web Parts page.
             */
            File.prototype.getLimitedWebPartManager = function (scope) {
                if (scope === void 0) { scope = WebPartsPersonalizationScope.Shared; }
                return new webparts_1.LimitedWebPartManager(this, "getLimitedWebPartManager(scope=" + scope + ")");
            };
            /**
             * Moves the file to the specified destination url.
             *
             * @param url The absolute url or server relative url of the destination file path to move to.
             * @param moveOperations The bitwise MoveOperations value for how to move the file.
             */
            File.prototype.moveTo = function (url, moveOperations) {
                if (moveOperations === void 0) { moveOperations = MoveOperations.Overwrite; }
                return new File(this, "moveTo(newurl='" + url + "',flags=" + moveOperations + ")").post();
            };
            /**
             * Submits the file for content approval with the specified comment.
             *
             * @param comment The comment for the published file. Its length must be <= 1023.
             */
            File.prototype.publish = function (comment) {
                if (comment === void 0) { comment = ""; }
                return new File(this, "publish(comment='" + comment + "')").post();
            };
            /**
             * Moves the file to the Recycle Bin and returns the identifier of the new Recycle Bin item.
             *
             * @returns The GUID of the recycled file.
             */
            File.prototype.recycle = function () {
                return new File(this, "recycle").post();
            };
            /**
             * Reverts an existing checkout for the file.
             *
             */
            File.prototype.undoCheckout = function () {
                return new File(this, "undoCheckout").post();
            };
            /**
             * Removes the file from content approval or unpublish a major version.
             *
             * @param comment The comment for the unpublish operation. Its length must be <= 1023.
             */
            File.prototype.unpublish = function (comment) {
                if (comment === void 0) { comment = ""; }
                if (comment.length > 1023) {
                    throw new exceptions_1.MaxCommentLengthException();
                }
                return new File(this, "unpublish(comment='" + comment + "')").post();
            };
            /**
             * Gets the contents of the file as text
             *
             */
            File.prototype.getText = function () {
                return new File(this, "$value").get(new odata_1.TextFileParser(), { headers: { "binaryStringResponseBody": "true" } });
            };
            /**
             * Gets the contents of the file as a blob, does not work in Node.js
             *
             */
            File.prototype.getBlob = function () {
                return new File(this, "$value").get(new odata_1.BlobFileParser(), { headers: { "binaryStringResponseBody": "true" } });
            };
            /**
             * Gets the contents of a file as an ArrayBuffer, works in Node.js
             */
            File.prototype.getBuffer = function () {
                return new File(this, "$value").get(new odata_1.BufferFileParser(), { headers: { "binaryStringResponseBody": "true" } });
            };
            /**
             * Gets the contents of a file as an ArrayBuffer, works in Node.js
             */
            File.prototype.getJSON = function () {
                return new File(this, "$value").get(new odata_1.JSONFileParser(), { headers: { "binaryStringResponseBody": "true" } });
            };
            /**
             * Sets the content of a file, for large files use setContentChunked
             *
             * @param content The file content
             *
             */
            File.prototype.setContent = function (content) {
                var _this = this;
                var setter = new File(this, "$value");
                return setter.post({
                    body: content,
                    headers: {
                        "X-HTTP-Method": "PUT",
                    },
                }).then(function (_) { return new File(_this); });
            };
            /**
             * Sets the contents of a file using a chunked upload approach
             *
             * @param file The file to upload
             * @param progress A callback function which can be used to track the progress of the upload
             * @param chunkSize The size of each file slice, in bytes (default: 10485760)
             */
            File.prototype.setContentChunked = function (file, progress, chunkSize) {
                if (chunkSize === void 0) { chunkSize = 10485760; }
                if (typeof progress === "undefined") {
                    progress = function () { return null; };
                }
                var self = this;
                var fileSize = file.size;
                var blockCount = parseInt((file.size / chunkSize).toString(), 10) + ((file.size % chunkSize === 0) ? 1 : 0);
                var uploadId = util_1.Util.getGUID();
                // start the chain with the first fragment
                progress({ blockNumber: 1, chunkSize: chunkSize, currentPointer: 0, fileSize: fileSize, stage: "starting", totalBlocks: blockCount });
                var chain = self.startUpload(uploadId, file.slice(0, chunkSize));
                var _loop_1 = function (i) {
                    chain = chain.then(function (pointer) {
                        progress({ blockNumber: i, chunkSize: chunkSize, currentPointer: pointer, fileSize: fileSize, stage: "continue", totalBlocks: blockCount });
                        return self.continueUpload(uploadId, pointer, file.slice(pointer, pointer + chunkSize));
                    });
                };
                // skip the first and last blocks
                for (var i = 2; i < blockCount; i++) {
                    _loop_1(i);
                }
                return chain.then(function (pointer) {
                    progress({ blockNumber: blockCount, chunkSize: chunkSize, currentPointer: pointer, fileSize: fileSize, stage: "finishing", totalBlocks: blockCount });
                    return self.finishUpload(uploadId, pointer, file.slice(pointer));
                }).then(function (_) {
                    return self;
                });
            };
            /**
             * Starts a new chunk upload session and uploads the first fragment.
             * The current file content is not changed when this method completes.
             * The method is idempotent (and therefore does not change the result) as long as you use the same values for uploadId and stream.
             * The upload session ends either when you use the CancelUpload method or when you successfully
             * complete the upload session by passing the rest of the file contents through the ContinueUpload and FinishUpload methods.
             * The StartUpload and ContinueUpload methods return the size of the running total of uploaded data in bytes,
             * so you can pass those return values to subsequent uses of ContinueUpload and FinishUpload.
             * This method is currently available only on Office 365.
             *
             * @param uploadId The unique identifier of the upload session.
             * @param fragment The file contents.
             * @returns The size of the total uploaded data in bytes.
             */
            File.prototype.startUpload = function (uploadId, fragment) {
                return new File(this, "startUpload(uploadId=guid'" + uploadId + "')").postAs({ body: fragment }).then(function (n) { return parseFloat(n); });
            };
            /**
             * Continues the chunk upload session with an additional fragment.
             * The current file content is not changed.
             * Use the uploadId value that was passed to the StartUpload method that started the upload session.
             * This method is currently available only on Office 365.
             *
             * @param uploadId The unique identifier of the upload session.
             * @param fileOffset The size of the offset into the file where the fragment starts.
             * @param fragment The file contents.
             * @returns The size of the total uploaded data in bytes.
             */
            File.prototype.continueUpload = function (uploadId, fileOffset, fragment) {
                return new File(this, "continueUpload(uploadId=guid'" + uploadId + "',fileOffset=" + fileOffset + ")").postAs({ body: fragment }).then(function (n) { return parseFloat(n); });
            };
            /**
             * Uploads the last file fragment and commits the file. The current file content is changed when this method completes.
             * Use the uploadId value that was passed to the StartUpload method that started the upload session.
             * This method is currently available only on Office 365.
             *
             * @param uploadId The unique identifier of the upload session.
             * @param fileOffset The size of the offset into the file where the fragment starts.
             * @param fragment The file contents.
             * @returns The newly uploaded file.
             */
            File.prototype.finishUpload = function (uploadId, fileOffset, fragment) {
                return new File(this, "finishUpload(uploadId=guid'" + uploadId + "',fileOffset=" + fileOffset + ")")
                    .postAs({ body: fragment }).then(function (response) {
                        return {
                            data: response,
                            file: new File(response.ServerRelativeUrl),
                        };
                    });
            };
            return File;
        }(queryable_1.QueryableInstance));
        exports.File = File;
        /**
         * Describes a collection of Version objects
         *
         */
        var Versions = (function (_super) {
            __extends(Versions, _super);
            /**
             * Creates a new instance of the File class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Versions(baseUrl, path) {
                if (path === void 0) { path = "versions"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a version by id
             *
             * @param versionId The id of the version to retrieve
             */
            Versions.prototype.getById = function (versionId) {
                var v = new Version(this);
                v.concat("(" + versionId + ")");
                return v;
            };
            /**
             * Deletes all the file version objects in the collection.
             *
             */
            Versions.prototype.deleteAll = function () {
                return new Versions(this, "deleteAll").post();
            };
            /**
             * Deletes the specified version of the file.
             *
             * @param versionId The ID of the file version to delete.
             */
            Versions.prototype.deleteById = function (versionId) {
                return new Versions(this, "deleteById(vid=" + versionId + ")").post();
            };
            /**
             * Deletes the file version object with the specified version label.
             *
             * @param label The version label of the file version to delete, for example: 1.2
             */
            Versions.prototype.deleteByLabel = function (label) {
                return new Versions(this, "deleteByLabel(versionlabel='" + label + "')").post();
            };
            /**
             * Creates a new file version from the file specified by the version label.
             *
             * @param label The version label of the file version to restore, for example: 1.2
             */
            Versions.prototype.restoreByLabel = function (label) {
                return new Versions(this, "restoreByLabel(versionlabel='" + label + "')").post();
            };
            return Versions;
        }(queryable_1.QueryableCollection));
        exports.Versions = Versions;
        /**
         * Describes a single Version instance
         *
         */
        var Version = (function (_super) {
            __extends(Version, _super);
            /**
             * Creates a new instance of the Version class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function Version(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
            * Delete a specific version of a file.
            *
            * @param eTag Value used in the IF-Match header, by default "*"
            */
            Version.prototype.delete = function (eTag) {
                if (eTag === void 0) { eTag = "*"; }
                return this.post({
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            return Version;
        }(queryable_1.QueryableInstance));
        exports.Version = Version;
        var CheckinType;
        (function (CheckinType) {
            CheckinType[CheckinType["Minor"] = 0] = "Minor";
            CheckinType[CheckinType["Major"] = 1] = "Major";
            CheckinType[CheckinType["Overwrite"] = 2] = "Overwrite";
        })(CheckinType = exports.CheckinType || (exports.CheckinType = {}));
        var WebPartsPersonalizationScope;
        (function (WebPartsPersonalizationScope) {
            WebPartsPersonalizationScope[WebPartsPersonalizationScope["User"] = 0] = "User";
            WebPartsPersonalizationScope[WebPartsPersonalizationScope["Shared"] = 1] = "Shared";
        })(WebPartsPersonalizationScope = exports.WebPartsPersonalizationScope || (exports.WebPartsPersonalizationScope = {}));
        var MoveOperations;
        (function (MoveOperations) {
            MoveOperations[MoveOperations["Overwrite"] = 1] = "Overwrite";
            MoveOperations[MoveOperations["AllowBrokenThickets"] = 8] = "AllowBrokenThickets";
        })(MoveOperations = exports.MoveOperations || (exports.MoveOperations = {}));
        var TemplateFileType;
        (function (TemplateFileType) {
            TemplateFileType[TemplateFileType["StandardPage"] = 0] = "StandardPage";
            TemplateFileType[TemplateFileType["WikiPage"] = 1] = "WikiPage";
            TemplateFileType[TemplateFileType["FormPage"] = 2] = "FormPage";
        })(TemplateFileType = exports.TemplateFileType || (exports.TemplateFileType = {}));


        /***/
    }),
    /* 8 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var items_1 = __webpack_require__(23);
        var views_1 = __webpack_require__(45);
        var contenttypes_1 = __webpack_require__(11);
        var fields_1 = __webpack_require__(21);
        var forms_1 = __webpack_require__(40);
        var subscriptions_1 = __webpack_require__(43);
        var queryable_1 = __webpack_require__(1);
        var queryablesecurable_1 = __webpack_require__(12);
        var util_1 = __webpack_require__(0);
        var usercustomactions_1 = __webpack_require__(15);
        var odata_1 = __webpack_require__(3);
        var exceptions_1 = __webpack_require__(2);
        /**
         * Describes a collection of List objects
         *
         */
        var Lists = (function (_super) {
            __extends(Lists, _super);
            /**
             * Creates a new instance of the Lists class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Lists(baseUrl, path) {
                if (path === void 0) { path = "lists"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a list from the collection by title
             *
             * @param title The title of the list
             */
            Lists.prototype.getByTitle = function (title) {
                return new List(this, "getByTitle('" + title + "')");
            };
            /**
             * Gets a list from the collection by guid id
             *
             * @param id The Id of the list (GUID)
             */
            Lists.prototype.getById = function (id) {
                var list = new List(this);
                list.concat("('" + id + "')");
                return list;
            };
            /**
             * Adds a new list to the collection
             *
             * @param title The new list's title
             * @param description The new list's description
             * @param template The list template value
             * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
             * @param additionalSettings Will be passed as part of the list creation body
             */
            Lists.prototype.add = function (title, description, template, enableContentTypes, additionalSettings) {
                var _this = this;
                if (description === void 0) { description = ""; }
                if (template === void 0) { template = 100; }
                if (enableContentTypes === void 0) { enableContentTypes = false; }
                if (additionalSettings === void 0) { additionalSettings = {}; }
                var addSettings = util_1.Util.extend({
                    "AllowContentTypes": enableContentTypes,
                    "BaseTemplate": template,
                    "ContentTypesEnabled": enableContentTypes,
                    "Description": description,
                    "Title": title,
                    "__metadata": { "type": "SP.List" },
                }, additionalSettings);
                return this.post({ body: JSON.stringify(addSettings) }).then(function (data) {
                    return { data: data, list: _this.getByTitle(addSettings.Title) };
                });
            };
            /**
             * Ensures that the specified list exists in the collection (note: this method not supported for batching)
             *
             * @param title The new list's title
             * @param description The new list's description
             * @param template The list template value
             * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
             * @param additionalSettings Will be passed as part of the list creation body or used to update an existing list
             */
            Lists.prototype.ensure = function (title, description, template, enableContentTypes, additionalSettings) {
                var _this = this;
                if (description === void 0) { description = ""; }
                if (template === void 0) { template = 100; }
                if (enableContentTypes === void 0) { enableContentTypes = false; }
                if (additionalSettings === void 0) { additionalSettings = {}; }
                if (this.hasBatch) {
                    throw new exceptions_1.NotSupportedInBatchException("The ensure list method");
                }
                return new Promise(function (resolve, reject) {
                    var addOrUpdateSettings = util_1.Util.extend(additionalSettings, { Title: title, Description: description, ContentTypesEnabled: enableContentTypes }, true);
                    var list = _this.getByTitle(addOrUpdateSettings.Title);
                    list.get().then(function (_) {
                        list.update(addOrUpdateSettings).then(function (d) {
                            resolve({ created: false, data: d, list: _this.getByTitle(addOrUpdateSettings.Title) });
                        }).catch(function (e) { return reject(e); });
                    }).catch(function (_) {
                        _this.add(title, description, template, enableContentTypes, addOrUpdateSettings).then(function (r) {
                            resolve({ created: true, data: r.data, list: _this.getByTitle(addOrUpdateSettings.Title) });
                        }).catch(function (e) { return reject(e); });
                    });
                });
            };
            /**
             * Gets a list that is the default asset location for images or other files, which the users upload to their wiki pages.
             */
            Lists.prototype.ensureSiteAssetsLibrary = function () {
                var q = new Lists(this, "ensuresiteassetslibrary");
                return q.post().then(function (json) {
                    return new List(odata_1.extractOdataId(json));
                });
            };
            /**
             * Gets a list that is the default location for wiki pages.
             */
            Lists.prototype.ensureSitePagesLibrary = function () {
                var q = new Lists(this, "ensuresitepageslibrary");
                return q.post().then(function (json) {
                    return new List(odata_1.extractOdataId(json));
                });
            };
            return Lists;
        }(queryable_1.QueryableCollection));
        exports.Lists = Lists;
        /**
         * Describes a single List instance
         *
         */
        var List = (function (_super) {
            __extends(List, _super);
            /**
             * Creates a new instance of the Lists class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function List(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(List.prototype, "contentTypes", {
                /**
                 * Gets the content types in this list
                 *
                 */
                get: function () {
                    return new contenttypes_1.ContentTypes(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "items", {
                /**
                 * Gets the items in this list
                 *
                 */
                get: function () {
                    return new items_1.Items(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "views", {
                /**
                 * Gets the views in this list
                 *
                 */
                get: function () {
                    return new views_1.Views(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "fields", {
                /**
                 * Gets the fields in this list
                 *
                 */
                get: function () {
                    return new fields_1.Fields(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "forms", {
                /**
                 * Gets the forms in this list
                 *
                 */
                get: function () {
                    return new forms_1.Forms(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "defaultView", {
                /**
                 * Gets the default view of this list
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableInstance(this, "DefaultView");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "userCustomActions", {
                /**
                 * Get all custom actions on a site collection
                 *
                 */
                get: function () {
                    return new usercustomactions_1.UserCustomActions(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "effectiveBasePermissions", {
                /**
                 * Gets the effective base permissions of this list
                 *
                 */
                get: function () {
                    return new queryable_1.Queryable(this, "EffectiveBasePermissions");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "eventReceivers", {
                /**
                 * Gets the event receivers attached to this list
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "EventReceivers");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "relatedFields", {
                /**
                 * Gets the related fields of this list
                 *
                 */
                get: function () {
                    return new queryable_1.Queryable(this, "getRelatedFields");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "informationRightsManagementSettings", {
                /**
                 * Gets the IRM settings for this list
                 *
                 */
                get: function () {
                    return new queryable_1.Queryable(this, "InformationRightsManagementSettings");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(List.prototype, "subscriptions", {
                /**
                 * Gets the webhook subscriptions of this list
                 *
                 */
                get: function () {
                    return new subscriptions_1.Subscriptions(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Gets a view by view guid id
             *
             */
            List.prototype.getView = function (viewId) {
                return new views_1.View(this, "getView('" + viewId + "')");
            };
            /**
             * Updates this list intance with the supplied properties
             *
             * @param properties A plain object hash of values to update for the list
             * @param eTag Value used in the IF-Match header, by default "*"
             */
            /* tslint:disable no-string-literal */
            List.prototype.update = function (properties, eTag) {
                var _this = this;
                if (eTag === void 0) { eTag = "*"; }
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": "SP.List" },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    var retList = _this;
                    if (properties.hasOwnProperty("Title")) {
                        retList = _this.getParent(List, _this.parentUrl, "getByTitle('" + properties["Title"] + "')");
                    }
                    return {
                        data: data,
                        list: retList,
                    };
                });
            };
            /* tslint:enable */
            /**
             * Delete this list
             *
             * @param eTag Value used in the IF-Match header, by default "*"
             */
            List.prototype.delete = function (eTag) {
                if (eTag === void 0) { eTag = "*"; }
                return this.post({
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            /**
             * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
             */
            List.prototype.getChanges = function (query) {
                var postBody = JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.ChangeQuery" } }, query) });
                // don't change "this" instance of the List, make a new one
                var q = new List(this, "getchanges");
                return q.post({ body: postBody });
            };
            /**
             * Returns a collection of items from the list based on the specified query.
             *
             * @param CamlQuery The Query schema of Collaborative Application Markup
             * Language (CAML) is used in various ways within the context of Microsoft SharePoint Foundation
             * to define queries against list data.
             * see:
             *
             * https://msdn.microsoft.com/en-us/library/office/ms467521.aspx
             *
             * @param expands A URI with a $expand System Query Option indicates that Entries associated with
             * the Entry or Collection of Entries identified by the Resource Path
             * section of the URI must be represented inline (i.e. eagerly loaded).
             * see:
             *
             * https://msdn.microsoft.com/en-us/library/office/fp142385.aspx
             *
             * http://www.odata.org/documentation/odata-version-2-0/uri-conventions/#ExpandSystemQueryOption
             */
            List.prototype.getItemsByCAMLQuery = function (query) {
                var expands = [];
                for (var _i = 1; _i < arguments.length; _i++) {
                    expands[_i - 1] = arguments[_i];
                }
                var postBody = JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.CamlQuery" } }, query) });
                // don't change "this" instance of the List, make a new one
                var q = new List(this, "getitems");
                return q.expand.apply(q, expands).post({ body: postBody });
            };
            /**
             * See: https://msdn.microsoft.com/en-us/library/office/dn292554.aspx
             */
            List.prototype.getListItemChangesSinceToken = function (query) {
                var postBody = JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.ChangeLogItemQuery" } }, query) });
                // don't change "this" instance of the List, make a new one
                var q = new List(this, "getlistitemchangessincetoken");
                // note we are using a custom parser to return text as the response is an xml doc
                return q.post({ body: postBody }, { parse: function (r) { return r.text(); } });
            };
            /**
             * Moves the list to the Recycle Bin and returns the identifier of the new Recycle Bin item.
             */
            List.prototype.recycle = function () {
                this.append("recycle");
                return this.post().then(function (data) {
                    if (data.hasOwnProperty("Recycle")) {
                        return data.Recycle;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Renders list data based on the view xml provided
             */
            List.prototype.renderListData = function (viewXml) {
                // don't change "this" instance of the List, make a new one
                var q = new List(this, "renderlistdata(@viewXml)");
                q.query.add("@viewXml", "'" + viewXml + "'");
                return q.post().then(function (data) {
                    // data will be a string, so we parse it again
                    data = JSON.parse(data);
                    if (data.hasOwnProperty("RenderListData")) {
                        return data.RenderListData;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Gets the field values and field schema attributes for a list item.
             */
            List.prototype.renderListFormData = function (itemId, formId, mode) {
                // don't change "this" instance of the List, make a new one
                var q = new List(this, "renderlistformdata(itemid=" + itemId + ", formid='" + formId + "', mode=" + mode + ")");
                return q.post().then(function (data) {
                    // data will be a string, so we parse it again
                    data = JSON.parse(data);
                    if (data.hasOwnProperty("ListData")) {
                        return data.ListData;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Reserves a list item ID for idempotent list item creation.
             */
            List.prototype.reserveListItemId = function () {
                // don't change "this" instance of the List, make a new one
                var q = new List(this, "reservelistitemid");
                return q.post().then(function (data) {
                    if (data.hasOwnProperty("ReserveListItemId")) {
                        return data.ReserveListItemId;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Returns the ListItemEntityTypeFullName for this list, used when adding/updating list items
             *
             */
            List.prototype.getListItemEntityTypeFullName = function () {
                var q = new queryable_1.QueryableInstance(this);
                return q.select("ListItemEntityTypeFullName").getAs().then(function (o) { return o.ListItemEntityTypeFullName; });
            };
            return List;
        }(queryablesecurable_1.QueryableSecurable));
        exports.List = List;


        /***/
    }),
    /* 9 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        /**
         * A wrapper class to provide a consistent interface to browser based storage
         *
         */
        var PnPClientStorageWrapper = (function () {
            /**
             * Creates a new instance of the PnPClientStorageWrapper class
             *
             * @constructor
             */
            function PnPClientStorageWrapper(store, defaultTimeoutMinutes) {
                this.store = store;
                this.defaultTimeoutMinutes = defaultTimeoutMinutes;
                this.defaultTimeoutMinutes = (defaultTimeoutMinutes === void 0) ? 5 : defaultTimeoutMinutes;
                this.enabled = this.test();
            }
            /**
             * Get a value from storage, or null if that value does not exist
             *
             * @param key The key whose value we want to retrieve
             */
            PnPClientStorageWrapper.prototype.get = function (key) {
                if (!this.enabled) {
                    return null;
                }
                var o = this.store.getItem(key);
                if (o == null) {
                    return null;
                }
                var persistable = JSON.parse(o);
                if (new Date(persistable.expiration) <= new Date()) {
                    this.delete(key);
                    return null;
                }
                else {
                    return persistable.value;
                }
            };
            /**
             * Adds a value to the underlying storage
             *
             * @param key The key to use when storing the provided value
             * @param o The value to store
             * @param expire Optional, if provided the expiration of the item, otherwise the default is used
             */
            PnPClientStorageWrapper.prototype.put = function (key, o, expire) {
                if (this.enabled) {
                    this.store.setItem(key, this.createPersistable(o, expire));
                }
            };
            /**
             * Deletes a value from the underlying storage
             *
             * @param key The key of the pair we want to remove from storage
             */
            PnPClientStorageWrapper.prototype.delete = function (key) {
                if (this.enabled) {
                    this.store.removeItem(key);
                }
            };
            /**
             * Gets an item from the underlying storage, or adds it if it does not exist using the supplied getter function
             *
             * @param key The key to use when storing the provided value
             * @param getter A function which will upon execution provide the desired value
             * @param expire Optional, if provided the expiration of the item, otherwise the default is used
             */
            PnPClientStorageWrapper.prototype.getOrPut = function (key, getter, expire) {
                var _this = this;
                if (!this.enabled) {
                    return getter();
                }
                return new Promise(function (resolve) {
                    var o = _this.get(key);
                    if (o == null) {
                        getter().then(function (d) {
                            _this.put(key, d, expire);
                            resolve(d);
                        });
                    }
                    else {
                        resolve(o);
                    }
                });
            };
            /**
             * Used to determine if the wrapped storage is available currently
             */
            PnPClientStorageWrapper.prototype.test = function () {
                var str = "test";
                try {
                    this.store.setItem(str, str);
                    this.store.removeItem(str);
                    return true;
                }
                catch (e) {
                    return false;
                }
            };
            /**
             * Creates the persistable to store
             */
            PnPClientStorageWrapper.prototype.createPersistable = function (o, expire) {
                if (typeof expire === "undefined") {
                    expire = util_1.Util.dateAdd(new Date(), "minute", this.defaultTimeoutMinutes);
                }
                return JSON.stringify({ expiration: expire, value: o });
            };
            return PnPClientStorageWrapper;
        }());
        exports.PnPClientStorageWrapper = PnPClientStorageWrapper;
        /**
         * A class that will establish wrappers for both local and session storage
         */
        var PnPClientStorage = (function () {
            /**
             * Creates a new instance of the PnPClientStorage class
             *
             * @constructor
             */
            function PnPClientStorage() {
                this.local = typeof localStorage !== "undefined" ? new PnPClientStorageWrapper(localStorage) : null;
                this.session = typeof sessionStorage !== "undefined" ? new PnPClientStorageWrapper(sessionStorage) : null;
            }
            return PnPClientStorage;
        }());
        exports.PnPClientStorage = PnPClientStorage;


        /***/
    }),
    /* 10 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var digestcache_1 = __webpack_require__(36);
        var util_1 = __webpack_require__(0);
        var pnplibconfig_1 = __webpack_require__(4);
        var exceptions_1 = __webpack_require__(2);
        var HttpClient = (function () {
            function HttpClient() {
                this._impl = pnplibconfig_1.RuntimeConfig.fetchClientFactory();
                this._digestCache = new digestcache_1.DigestCache(this);
            }
            HttpClient.prototype.fetch = function (url, options) {
                var _this = this;
                if (options === void 0) { options = {}; }
                var opts = util_1.Util.extend(options, { cache: "no-cache", credentials: "same-origin" }, true);
                var headers = new Headers();
                // first we add the global headers so they can be overwritten by any passed in locally to this call
                this.mergeHeaders(headers, pnplibconfig_1.RuntimeConfig.headers);
                // second we add the local options so we can overwrite the globals
                this.mergeHeaders(headers, options.headers);
                // lastly we apply any default headers we need that may not exist
                if (!headers.has("Accept")) {
                    headers.append("Accept", "application/json");
                }
                if (!headers.has("Content-Type")) {
                    headers.append("Content-Type", "application/json;odata=verbose;charset=utf-8");
                }
                if (!headers.has("X-ClientService-ClientTag")) {
                    headers.append("X-ClientService-ClientTag", "PnPCoreJS:2.0.2");
                }
                opts = util_1.Util.extend(opts, { headers: headers });
                if (opts.method && opts.method.toUpperCase() !== "GET") {
                    if (!headers.has("X-RequestDigest")) {
                        var index = url.indexOf("_api/");
                        if (index < 0) {
                            throw new exceptions_1.APIUrlException();
                        }
                        var webUrl = url.substr(0, index);
                        return this._digestCache.getDigest(webUrl)
                            .then(function (digest) {
                                headers.append("X-RequestDigest", digest);
                                return _this.fetchRaw(url, opts);
                            });
                    }
                }
                return this.fetchRaw(url, opts);
            };
            HttpClient.prototype.fetchRaw = function (url, options) {
                var _this = this;
                if (options === void 0) { options = {}; }
                // here we need to normalize the headers
                var rawHeaders = new Headers();
                this.mergeHeaders(rawHeaders, options.headers);
                options = util_1.Util.extend(options, { headers: rawHeaders });
                var retry = function (ctx) {
                    _this._impl.fetch(url, options).then(function (response) { return ctx.resolve(response); }).catch(function (response) {
                        // grab our current delay
                        var delay = ctx.delay;
                        // Check if request was throttled - http status code 429
                        // Check is request failed due to server unavailable - http status code 503
                        if (response.status !== 429 && response.status !== 503) {
                            ctx.reject(response);
                        }
                        // Increment our counters.
                        ctx.delay *= 2;
                        ctx.attempts++;
                        // If we have exceeded the retry count, reject.
                        if (ctx.retryCount <= ctx.attempts) {
                            ctx.reject(response);
                        }
                        // Set our retry timeout for {delay} milliseconds.
                        setTimeout(util_1.Util.getCtxCallback(_this, retry, ctx), delay);
                    });
                };
                return new Promise(function (resolve, reject) {
                    var retryContext = {
                        attempts: 0,
                        delay: 100,
                        reject: reject,
                        resolve: resolve,
                        retryCount: 7,
                    };
                    retry.call(_this, retryContext);
                });
            };
            HttpClient.prototype.get = function (url, options) {
                if (options === void 0) { options = {}; }
                var opts = util_1.Util.extend(options, { method: "GET" });
                return this.fetch(url, opts);
            };
            HttpClient.prototype.post = function (url, options) {
                if (options === void 0) { options = {}; }
                var opts = util_1.Util.extend(options, { method: "POST" });
                return this.fetch(url, opts);
            };
            HttpClient.prototype.patch = function (url, options) {
                if (options === void 0) { options = {}; }
                var opts = util_1.Util.extend(options, { method: "PATCH" });
                return this.fetch(url, opts);
            };
            HttpClient.prototype.delete = function (url, options) {
                if (options === void 0) { options = {}; }
                var opts = util_1.Util.extend(options, { method: "DELETE" });
                return this.fetch(url, opts);
            };
            HttpClient.prototype.mergeHeaders = function (target, source) {
                if (typeof source !== "undefined" && source !== null) {
                    var temp = new Request("", { headers: source });
                    temp.headers.forEach(function (value, name) {
                        target.append(name, value);
                    });
                }
            };
            return HttpClient;
        }());
        exports.HttpClient = HttpClient;
        ;


        /***/
    }),
    /* 11 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        var queryable_1 = __webpack_require__(1);
        /**
         * Describes a collection of content types
         *
         */
        var ContentTypes = (function (_super) {
            __extends(ContentTypes, _super);
            /**
             * Creates a new instance of the ContentTypes class
             *
             * @param baseUrl The url or Queryable which forms the parent of this content types collection
             */
            function ContentTypes(baseUrl, path) {
                if (path === void 0) { path = "contenttypes"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a ContentType by content type id
             */
            ContentTypes.prototype.getById = function (id) {
                var ct = new ContentType(this);
                ct.concat("('" + id + "')");
                return ct;
            };
            /**
             * Adds an existing contenttype to a content type collection
             *
             * @param contentTypeId in the following format, for example: 0x010102
             */
            ContentTypes.prototype.addAvailableContentType = function (contentTypeId) {
                var _this = this;
                var postBody = JSON.stringify({
                    "contentTypeId": contentTypeId,
                });
                return new ContentTypes(this, "addAvailableContentType").postAs({ body: postBody }).then(function (data) {
                    return {
                        contentType: _this.getById(data.id),
                        data: data,
                    };
                });
            };
            /**
             * Adds a new content type to the collection
             *
             * @param id The desired content type id for the new content type (also determines the parent content type)
             * @param name The name of the content type
             * @param description The description of the content type
             * @param group The group in which to add the content type
             * @param additionalSettings Any additional settings to provide when creating the content type
             *
             */
            ContentTypes.prototype.add = function (id, name, description, group, additionalSettings) {
                var _this = this;
                if (description === void 0) { description = ""; }
                if (group === void 0) { group = "Custom Content Types"; }
                if (additionalSettings === void 0) { additionalSettings = {}; }
                var postBody = JSON.stringify(util_1.Util.extend({
                    "Description": description,
                    "Group": group,
                    "Id": { "StringValue": id },
                    "Name": name,
                    "__metadata": { "type": "SP.ContentType" },
                }, additionalSettings));
                return this.post({ body: postBody }).then(function (data) {
                    return { contentType: _this.getById(data.id), data: data };
                });
            };
            return ContentTypes;
        }(queryable_1.QueryableCollection));
        exports.ContentTypes = ContentTypes;
        /**
         * Describes a single ContentType instance
         *
         */
        var ContentType = (function (_super) {
            __extends(ContentType, _super);
            /**
             * Creates a new instance of the ContentType class
             *
             * @param baseUrl The url or Queryable which forms the parent of this content type instance
             */
            function ContentType(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(ContentType.prototype, "fieldLinks", {
                /**
                 * Gets the column (also known as field) references in the content type.
                */
                get: function () {
                    return new FieldLinks(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ContentType.prototype, "fields", {
                /**
                 * Gets a value that specifies the collection of fields for the content type.
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "fields");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ContentType.prototype, "parent", {
                /**
                 * Gets the parent content type of the content type.
                 */
                get: function () {
                    return new ContentType(this, "parent");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ContentType.prototype, "workflowAssociations", {
                /**
                 * Gets a value that specifies the collection of workflow associations for the content type.
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "workflowAssociations");
                },
                enumerable: true,
                configurable: true
            });
            return ContentType;
        }(queryable_1.QueryableInstance));
        exports.ContentType = ContentType;
        /**
         * Represents a collection of field link instances
         */
        var FieldLinks = (function (_super) {
            __extends(FieldLinks, _super);
            /**
             * Creates a new instance of the ContentType class
             *
             * @param baseUrl The url or Queryable which forms the parent of this content type instance
             */
            function FieldLinks(baseUrl, path) {
                if (path === void 0) { path = "fieldlinks"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a FieldLink by GUID id
             *
             * @param id The GUID id of the field link
             */
            FieldLinks.prototype.getById = function (id) {
                var fl = new FieldLink(this);
                fl.concat("(guid'" + id + "')");
                return fl;
            };
            return FieldLinks;
        }(queryable_1.QueryableCollection));
        exports.FieldLinks = FieldLinks;
        /**
         * Represents a field link instance
         */
        var FieldLink = (function (_super) {
            __extends(FieldLink, _super);
            /**
             * Creates a new instance of the ContentType class
            *
            * @param baseUrl The url or Queryable which forms the parent of this content type instance
            */
            function FieldLink(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            return FieldLink;
        }(queryable_1.QueryableInstance));
        exports.FieldLink = FieldLink;


        /***/
    }),
    /* 12 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var roles_1 = __webpack_require__(13);
        var queryable_1 = __webpack_require__(1);
        var QueryableSecurable = (function (_super) {
            __extends(QueryableSecurable, _super);
            function QueryableSecurable() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            Object.defineProperty(QueryableSecurable.prototype, "roleAssignments", {
                /**
                 * Gets the set of role assignments for this item
                 *
                 */
                get: function () {
                    return new roles_1.RoleAssignments(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(QueryableSecurable.prototype, "firstUniqueAncestorSecurableObject", {
                /**
                 * Gets the closest securable up the security hierarchy whose permissions are applied to this list item
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableInstance(this, "FirstUniqueAncestorSecurableObject");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Gets the effective permissions for the user supplied
             *
             * @param loginName The claims username for the user (ex: i:0#.f|membership|user@domain.com)
             */
            QueryableSecurable.prototype.getUserEffectivePermissions = function (loginName) {
                var perms = new queryable_1.Queryable(this, "getUserEffectivePermissions(@user)");
                perms.query.add("@user", "'" + encodeURIComponent(loginName) + "'");
                return perms;
            };
            /**
             * Breaks the security inheritance at this level optinally copying permissions and clearing subscopes
             *
             * @param copyRoleAssignments If true the permissions are copied from the current parent scope
             * @param clearSubscopes Optional. true to make all child securable objects inherit role assignments from the current object
             */
            QueryableSecurable.prototype.breakRoleInheritance = function (copyRoleAssignments, clearSubscopes) {
                if (copyRoleAssignments === void 0) { copyRoleAssignments = false; }
                if (clearSubscopes === void 0) { clearSubscopes = false; }
                var Breaker = (function (_super) {
                    __extends(Breaker, _super);
                    function Breaker(baseUrl, copy, clear) {
                        return _super.call(this, baseUrl, "breakroleinheritance(copyroleassignments=" + copy + ", clearsubscopes=" + clear + ")") || this;
                    }
                    Breaker.prototype.break = function () {
                        return this.post();
                    };
                    return Breaker;
                }(queryable_1.Queryable));
                var b = new Breaker(this, copyRoleAssignments, clearSubscopes);
                return b.break();
            };
            /**
             * Removes the local role assignments so that it re-inherit role assignments from the parent object.
             *
             */
            QueryableSecurable.prototype.resetRoleInheritance = function () {
                var Resetter = (function (_super) {
                    __extends(Resetter, _super);
                    function Resetter(baseUrl) {
                        return _super.call(this, baseUrl, "resetroleinheritance") || this;
                    }
                    Resetter.prototype.reset = function () {
                        return this.post();
                    };
                    return Resetter;
                }(queryable_1.Queryable));
                var r = new Resetter(this);
                return r.reset();
            };
            return QueryableSecurable;
        }(queryable_1.QueryableInstance));
        exports.QueryableSecurable = QueryableSecurable;


        /***/
    }),
    /* 13 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var sitegroups_1 = __webpack_require__(14);
        var util_1 = __webpack_require__(0);
        /**
         * Describes a set of role assignments for the current scope
         *
         */
        var RoleAssignments = (function (_super) {
            __extends(RoleAssignments, _super);
            /**
             * Creates a new instance of the RoleAssignments class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function RoleAssignments(baseUrl, path) {
                if (path === void 0) { path = "roleassignments"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Adds a new role assignment with the specified principal and role definitions to the collection.
             *
             * @param principalId The ID of the user or group to assign permissions to
             * @param roleDefId The ID of the role definition that defines the permissions to assign
             *
             */
            RoleAssignments.prototype.add = function (principalId, roleDefId) {
                var a = new RoleAssignments(this, "addroleassignment(principalid=" + principalId + ", roledefid=" + roleDefId + ")");
                return a.post();
            };
            /**
             * Removes the role assignment with the specified principal and role definition from the collection
             *
             * @param principalId The ID of the user or group in the role assignment.
             * @param roleDefId The ID of the role definition in the role assignment
             *
             */
            RoleAssignments.prototype.remove = function (principalId, roleDefId) {
                var a = new RoleAssignments(this, "removeroleassignment(principalid=" + principalId + ", roledefid=" + roleDefId + ")");
                return a.post();
            };
            /**
             * Gets the role assignment associated with the specified principal ID from the collection.
             *
             * @param id The id of the role assignment
             */
            RoleAssignments.prototype.getById = function (id) {
                var ra = new RoleAssignment(this);
                ra.concat("(" + id + ")");
                return ra;
            };
            return RoleAssignments;
        }(queryable_1.QueryableCollection));
        exports.RoleAssignments = RoleAssignments;
        var RoleAssignment = (function (_super) {
            __extends(RoleAssignment, _super);
            /**
         * Creates a new instance of the RoleAssignment class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
            function RoleAssignment(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(RoleAssignment.prototype, "groups", {
                get: function () {
                    return new sitegroups_1.SiteGroups(this, "groups");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RoleAssignment.prototype, "bindings", {
                /**
                 * Get the role definition bindings for this role assignment
                 *
                 */
                get: function () {
                    return new RoleDefinitionBindings(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Delete this role assignment
             *
             */
            RoleAssignment.prototype.delete = function () {
                return this.post({
                    headers: {
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            return RoleAssignment;
        }(queryable_1.QueryableInstance));
        exports.RoleAssignment = RoleAssignment;
        var RoleDefinitions = (function (_super) {
            __extends(RoleDefinitions, _super);
            /**
             * Creates a new instance of the RoleDefinitions class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path
             *
             */
            function RoleDefinitions(baseUrl, path) {
                if (path === void 0) { path = "roledefinitions"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets the role definition with the specified ID from the collection.
             *
             * @param id The ID of the role definition.
             *
             */
            RoleDefinitions.prototype.getById = function (id) {
                return new RoleDefinition(this, "getById(" + id + ")");
            };
            /**
             * Gets the role definition with the specified name.
             *
             * @param name The name of the role definition.
             *
             */
            RoleDefinitions.prototype.getByName = function (name) {
                return new RoleDefinition(this, "getbyname('" + name + "')");
            };
            /**
             * Gets the role definition with the specified type.
             *
             * @param name The name of the role definition.
             *
             */
            RoleDefinitions.prototype.getByType = function (roleTypeKind) {
                return new RoleDefinition(this, "getbytype(" + roleTypeKind + ")");
            };
            /**
             * Create a role definition
             *
             * @param name The new role definition's name
             * @param description The new role definition's description
             * @param order The order in which the role definition appears
             * @param basePermissions The permissions mask for this role definition
             *
             */
            RoleDefinitions.prototype.add = function (name, description, order, basePermissions) {
                var _this = this;
                var postBody = JSON.stringify({
                    BasePermissions: util_1.Util.extend({ __metadata: { type: "SP.BasePermissions" } }, basePermissions),
                    Description: description,
                    Name: name,
                    Order: order,
                    __metadata: { "type": "SP.RoleDefinition" },
                });
                return this.post({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        definition: _this.getById(data.Id),
                    };
                });
            };
            return RoleDefinitions;
        }(queryable_1.QueryableCollection));
        exports.RoleDefinitions = RoleDefinitions;
        var RoleDefinition = (function (_super) {
            __extends(RoleDefinition, _super);
            function RoleDefinition(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Updates this web intance with the supplied properties
             *
             * @param properties A plain object hash of values to update for the web
             */
            /* tslint:disable no-string-literal */
            RoleDefinition.prototype.update = function (properties) {
                var _this = this;
                if (typeof properties.hasOwnProperty("BasePermissions") !== "undefined") {
                    properties["BasePermissions"] = util_1.Util.extend({ __metadata: { type: "SP.BasePermissions" } }, properties["BasePermissions"]);
                }
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": "SP.RoleDefinition" },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    var retDef = _this;
                    if (properties.hasOwnProperty("Name")) {
                        var parent_1 = _this.getParent(RoleDefinitions, _this.parentUrl, "");
                        retDef = parent_1.getByName(properties["Name"]);
                    }
                    return {
                        data: data,
                        definition: retDef,
                    };
                });
            };
            /* tslint:enable */
            /**
             * Delete this role definition
             *
             */
            RoleDefinition.prototype.delete = function () {
                return this.post({
                    headers: {
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            return RoleDefinition;
        }(queryable_1.QueryableInstance));
        exports.RoleDefinition = RoleDefinition;
        var RoleDefinitionBindings = (function (_super) {
            __extends(RoleDefinitionBindings, _super);
            function RoleDefinitionBindings(baseUrl, path) {
                if (path === void 0) { path = "roledefinitionbindings"; }
                return _super.call(this, baseUrl, path) || this;
            }
            return RoleDefinitionBindings;
        }(queryable_1.QueryableCollection));
        exports.RoleDefinitionBindings = RoleDefinitionBindings;


        /***/
    }),
    /* 14 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var siteusers_1 = __webpack_require__(28);
        var util_1 = __webpack_require__(0);
        /**
         * Principal Type enum
         *
         */
        var PrincipalType;
        (function (PrincipalType) {
            PrincipalType[PrincipalType["None"] = 0] = "None";
            PrincipalType[PrincipalType["User"] = 1] = "User";
            PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
            PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
            PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
            PrincipalType[PrincipalType["All"] = 15] = "All";
        })(PrincipalType = exports.PrincipalType || (exports.PrincipalType = {}));
        /**
         * Describes a collection of site users
         *
         */
        var SiteGroups = (function (_super) {
            __extends(SiteGroups, _super);
            /**
             * Creates a new instance of the SiteUsers class
             *
             * @param baseUrl The url or Queryable which forms the parent of this user collection
             */
            function SiteGroups(baseUrl, path) {
                if (path === void 0) { path = "sitegroups"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Adds a new group to the site collection
             *
             * @param props The properties to be updated
             */
            SiteGroups.prototype.add = function (properties) {
                var _this = this;
                var postBody = JSON.stringify(util_1.Util.extend({ "__metadata": { "type": "SP.Group" } }, properties));
                return this.post({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        group: _this.getById(data.Id),
                    };
                });
            };
            /**
             * Gets a group from the collection by name
             *
             * @param email The name of the group
             */
            SiteGroups.prototype.getByName = function (groupName) {
                return new SiteGroup(this, "getByName('" + groupName + "')");
            };
            /**
             * Gets a group from the collection by id
             *
             * @param id The id of the group
             */
            SiteGroups.prototype.getById = function (id) {
                var sg = new SiteGroup(this);
                sg.concat("(" + id + ")");
                return sg;
            };
            /**
             * Removes the group with the specified member ID from the collection.
             *
             * @param id The id of the group to remove
             */
            SiteGroups.prototype.removeById = function (id) {
                var g = new SiteGroups(this, "removeById('" + id + "')");
                return g.post();
            };
            /**
             * Removes a user from the collection by login name
             *
             * @param loginName The login name of the user
             */
            SiteGroups.prototype.removeByLoginName = function (loginName) {
                var g = new SiteGroups(this, "removeByLoginName('" + loginName + "')");
                return g.post();
            };
            return SiteGroups;
        }(queryable_1.QueryableCollection));
        exports.SiteGroups = SiteGroups;
        /**
         * Describes a single group
         *
         */
        var SiteGroup = (function (_super) {
            __extends(SiteGroup, _super);
            /**
             * Creates a new instance of the Group class
             *
             * @param baseUrl The url or Queryable which forms the parent of this site group
             * @param path Optional, passes the path to the group
             */
            function SiteGroup(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(SiteGroup.prototype, "users", {
                /**
                 * Get's the users for this group
                 *
                 */
                get: function () {
                    return new siteusers_1.SiteUsers(this, "users");
                },
                enumerable: true,
                configurable: true
            });
            /**
            * Updates this group instance with the supplied properties
            *
            * @param properties A GroupWriteableProperties object of property names and values to update for the user
            */
            /* tslint:disable no-string-literal */
            SiteGroup.prototype.update = function (properties) {
                var _this = this;
                var postBody = util_1.Util.extend({ "__metadata": { "type": "SP.Group" } }, properties);
                return this.post({
                    body: JSON.stringify(postBody),
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    var retGroup = _this;
                    if (properties.hasOwnProperty("Title")) {
                        retGroup = _this.getParent(SiteGroup, _this.parentUrl, "getByName('" + properties["Title"] + "')");
                    }
                    return {
                        data: data,
                        group: retGroup,
                    };
                });
            };
            return SiteGroup;
        }(queryable_1.QueryableInstance));
        exports.SiteGroup = SiteGroup;


        /***/
    }),
    /* 15 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var util_1 = __webpack_require__(0);
        var UserCustomActions = (function (_super) {
            __extends(UserCustomActions, _super);
            function UserCustomActions(baseUrl, path) {
                if (path === void 0) { path = "usercustomactions"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Returns the custom action with the specified identifier.
             *
             * @param id The GUID ID of the user custom action to get.
             */
            UserCustomActions.prototype.getById = function (id) {
                var uca = new UserCustomAction(this);
                uca.concat("('" + id + "')");
                return uca;
            };
            /**
             * Create a custom action
             *
             * @param creationInfo The information which defines the new custom action
             *
             */
            UserCustomActions.prototype.add = function (properties) {
                var _this = this;
                var postBody = JSON.stringify(util_1.Util.extend({ __metadata: { "type": "SP.UserCustomAction" } }, properties));
                return this.post({ body: postBody }).then(function (data) {
                    return {
                        action: _this.getById(data.Id),
                        data: data,
                    };
                });
            };
            /**
             * Deletes all custom actions in the collection.
             *
             */
            UserCustomActions.prototype.clear = function () {
                var a = new UserCustomActions(this, "clear");
                return a.post();
            };
            return UserCustomActions;
        }(queryable_1.QueryableCollection));
        exports.UserCustomActions = UserCustomActions;
        var UserCustomAction = (function (_super) {
            __extends(UserCustomAction, _super);
            function UserCustomAction(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            UserCustomAction.prototype.update = function (properties) {
                var _this = this;
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": "SP.UserCustomAction" },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    return {
                        action: _this,
                        data: data,
                    };
                });
            };
            /**
            * Remove a custom action
            *
            */
            UserCustomAction.prototype.delete = function () {
                return _super.prototype.delete.call(this);
            };
            return UserCustomAction;
        }(queryable_1.QueryableInstance));
        exports.UserCustomAction = UserCustomAction;


        /***/
    }),
    /* 16 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var queryablesecurable_1 = __webpack_require__(12);
        var lists_1 = __webpack_require__(8);
        var fields_1 = __webpack_require__(21);
        var navigation_1 = __webpack_require__(24);
        var sitegroups_1 = __webpack_require__(14);
        var contenttypes_1 = __webpack_require__(11);
        var folders_1 = __webpack_require__(22);
        var roles_1 = __webpack_require__(13);
        var files_1 = __webpack_require__(7);
        var util_1 = __webpack_require__(0);
        var lists_2 = __webpack_require__(8);
        var siteusers_1 = __webpack_require__(28);
        var usercustomactions_1 = __webpack_require__(15);
        var odata_1 = __webpack_require__(3);
        var features_1 = __webpack_require__(20);
        var Webs = (function (_super) {
            __extends(Webs, _super);
            function Webs(baseUrl, webPath) {
                if (webPath === void 0) { webPath = "webs"; }
                return _super.call(this, baseUrl, webPath) || this;
            }
            /**
             * Adds a new web to the collection
             *
             * @param title The new web's title
             * @param url The new web's relative url
             * @param description The web web's description
             * @param template The web's template
             * @param language The language code to use for this web
             * @param inheritPermissions If true permissions will be inherited from the partent web
             * @param additionalSettings Will be passed as part of the web creation body
             */
            Webs.prototype.add = function (title, url, description, template, language, inheritPermissions, additionalSettings) {
                if (description === void 0) { description = ""; }
                if (template === void 0) { template = "STS"; }
                if (language === void 0) { language = 1033; }
                if (inheritPermissions === void 0) { inheritPermissions = true; }
                if (additionalSettings === void 0) { additionalSettings = {}; }
                var props = util_1.Util.extend({
                    Description: description,
                    Language: language,
                    Title: title,
                    Url: url,
                    UseSamePermissionsAsParentSite: inheritPermissions,
                    WebTemplate: template,
                }, additionalSettings);
                var postBody = JSON.stringify({
                    "parameters": util_1.Util.extend({
                        "__metadata": { "type": "SP.WebCreationInformation" },
                    }, props),
                });
                var q = new Webs(this, "add");
                return q.post({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        web: new Web(odata_1.extractOdataId(data).replace(/_api\/web\/?/i, "")),
                    };
                });
            };
            return Webs;
        }(queryable_1.QueryableCollection));
        exports.Webs = Webs;
        /**
         * Describes a web
         *
         */
        var Web = (function (_super) {
            __extends(Web, _super);
            function Web(baseUrl, path) {
                if (path === void 0) { path = "_api/web"; }
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(Web.prototype, "webs", {
                get: function () {
                    return new Webs(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "contentTypes", {
                /**
                 * Get the content types available in this web
                 *
                 */
                get: function () {
                    return new contenttypes_1.ContentTypes(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "lists", {
                /**
                 * Get the lists in this web
                 *
                 */
                get: function () {
                    return new lists_1.Lists(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "fields", {
                /**
                 * Gets the fields in this web
                 *
                 */
                get: function () {
                    return new fields_1.Fields(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "features", {
                /**
                 * Gets the active features for this web
                 *
                 */
                get: function () {
                    return new features_1.Features(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "availablefields", {
                /**
                 * Gets the available fields in this web
                 *
                 */
                get: function () {
                    return new fields_1.Fields(this, "availablefields");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "navigation", {
                /**
                 * Get the navigation options in this web
                 *
                 */
                get: function () {
                    return new navigation_1.Navigation(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "siteUsers", {
                /**
                 * Gets the site users
                 *
                 */
                get: function () {
                    return new siteusers_1.SiteUsers(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "siteGroups", {
                /**
                 * Gets the site groups
                 *
                 */
                get: function () {
                    return new sitegroups_1.SiteGroups(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "currentUser", {
                /**
                 * Gets the current user
                 */
                get: function () {
                    return new siteusers_1.CurrentUser(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "folders", {
                /**
                 * Get the folders in this web
                 *
                 */
                get: function () {
                    return new folders_1.Folders(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "userCustomActions", {
                /**
                 * Get all custom actions on a site
                 *
                 */
                get: function () {
                    return new usercustomactions_1.UserCustomActions(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Web.prototype, "roleDefinitions", {
                /**
                 * Gets the collection of RoleDefinition resources.
                 *
                 */
                get: function () {
                    return new roles_1.RoleDefinitions(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Creates a new batch for requests within the context of context this web
             *
             */
            Web.prototype.createBatch = function () {
                return new odata_1.ODataBatch(this.parentUrl);
            };
            /**
             * Get a folder by server relative url
             *
             * @param folderRelativeUrl the server relative path to the folder (including /sites/ if applicable)
             */
            Web.prototype.getFolderByServerRelativeUrl = function (folderRelativeUrl) {
                return new folders_1.Folder(this, "getFolderByServerRelativeUrl('" + folderRelativeUrl + "')");
            };
            /**
             * Get a file by server relative url
             *
             * @param fileRelativeUrl the server relative path to the file (including /sites/ if applicable)
             */
            Web.prototype.getFileByServerRelativeUrl = function (fileRelativeUrl) {
                return new files_1.File(this, "getFileByServerRelativeUrl('" + fileRelativeUrl + "')");
            };
            /**
             * Get a list by server relative url (list's root folder)
             *
             * @param listRelativeUrl the server relative path to the list's root folder (including /sites/ if applicable)
             */
            Web.prototype.getList = function (listRelativeUrl) {
                return new lists_2.List(this, "getList('" + listRelativeUrl + "')");
            };
            /**
             * Updates this web intance with the supplied properties
             *
             * @param properties A plain object hash of values to update for the web
             */
            Web.prototype.update = function (properties) {
                var _this = this;
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": "SP.Web" },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    return {
                        data: data,
                        web: _this,
                    };
                });
            };
            /**
             * Delete this web
             *
             */
            Web.prototype.delete = function () {
                return _super.prototype.delete.call(this);
            };
            /**
             * Applies the theme specified by the contents of each of the files specified in the arguments to the site.
             *
             * @param colorPaletteUrl Server-relative URL of the color palette file.
             * @param fontSchemeUrl Server-relative URL of the font scheme.
             * @param backgroundImageUrl Server-relative URL of the background image.
             * @param shareGenerated true to store the generated theme files in the root site, or false to store them in this site.
             */
            Web.prototype.applyTheme = function (colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated) {
                var postBody = JSON.stringify({
                    backgroundImageUrl: backgroundImageUrl,
                    colorPaletteUrl: colorPaletteUrl,
                    fontSchemeUrl: fontSchemeUrl,
                    shareGenerated: shareGenerated,
                });
                var q = new Web(this, "applytheme");
                return q.post({ body: postBody });
            };
            /**
             * Applies the specified site definition or site template to the Web site that has no template applied to it.
             *
             * @param template Name of the site definition or the name of the site template
             */
            Web.prototype.applyWebTemplate = function (template) {
                var q = new Web(this, "applywebtemplate");
                q.concat("(@t)");
                q.query.add("@t", template);
                return q.post();
            };
            /**
             * Returns whether the current user has the given set of permissions.
             *
             * @param perms The high and low permission range.
             */
            Web.prototype.doesUserHavePermissions = function (perms) {
                var q = new Web(this, "doesuserhavepermissions");
                q.concat("(@p)");
                q.query.add("@p", JSON.stringify(perms));
                return q.get();
            };
            /**
             * Checks whether the specified login name belongs to a valid user in the site. If the user doesn't exist, adds the user to the site.
             *
             * @param loginName The login name of the user (ex: i:0#.f|membership|user@domain.onmicrosoft.com)
             */
            Web.prototype.ensureUser = function (loginName) {
                // TODO:: this should resolve to a User
                var postBody = JSON.stringify({
                    logonName: loginName,
                });
                var q = new Web(this, "ensureuser");
                return q.post({ body: postBody });
            };
            /**
             * Returns a collection of site templates available for the site.
             *
             * @param language The LCID of the site templates to get.
             * @param true to include language-neutral site templates; otherwise false
             */
            Web.prototype.availableWebTemplates = function (language, includeCrossLanugage) {
                if (language === void 0) { language = 1033; }
                if (includeCrossLanugage === void 0) { includeCrossLanugage = true; }
                return new queryable_1.QueryableCollection(this, "getavailablewebtemplates(lcid=" + language + ", doincludecrosslanguage=" + includeCrossLanugage + ")");
            };
            /**
             * Returns the list gallery on the site.
             *
             * @param type The gallery type - WebTemplateCatalog = 111, WebPartCatalog = 113 ListTemplateCatalog = 114,
             * MasterPageCatalog = 116, SolutionCatalog = 121, ThemeCatalog = 123, DesignCatalog = 124, AppDataCatalog = 125
             */
            Web.prototype.getCatalog = function (type) {
                var q = new Web(this, "getcatalog(" + type + ")");
                q.select("Id");
                return q.get().then(function (data) {
                    return new lists_2.List(odata_1.extractOdataId(data));
                });
            };
            /**
             * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
             */
            Web.prototype.getChanges = function (query) {
                var postBody = JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.ChangeQuery" } }, query) });
                // don't change "this" instance, make a new one
                var q = new Web(this, "getchanges");
                return q.post({ body: postBody });
            };
            Object.defineProperty(Web.prototype, "customListTemplate", {
                /**
                 * Gets the custom list templates for the site.
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "getcustomlisttemplates");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Returns the user corresponding to the specified member identifier for the current site.
             *
             * @param id The ID of the user.
             */
            Web.prototype.getUserById = function (id) {
                return new siteusers_1.SiteUser(this, "getUserById(" + id + ")");
            };
            /**
             * Returns the name of the image file for the icon that is used to represent the specified file.
             *
             * @param filename The file name. If this parameter is empty, the server returns an empty string.
             * @param size The size of the icon: 16x16 pixels = 0, 32x32 pixels = 1.
             * @param progId The ProgID of the application that was used to create the file, in the form OLEServerName.ObjectName
             */
            Web.prototype.mapToIcon = function (filename, size, progId) {
                if (size === void 0) { size = 0; }
                if (progId === void 0) { progId = ""; }
                var q = new Web(this, "maptoicon(filename='" + filename + "', progid='" + progId + "', size=" + size + ")");
                return q.get();
            };
            return Web;
        }(queryablesecurable_1.QueryableSecurable));
        exports.Web = Web;


        /***/
    }),
    /* 17 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var storage = __webpack_require__(9);
        var exceptions_1 = __webpack_require__(2);
        /**
         * A caching provider which can wrap other non-caching providers
         *
         */
        var CachingConfigurationProvider = (function () {
            /**
             * Creates a new caching configuration provider
             * @constructor
             * @param {IConfigurationProvider} wrappedProvider Provider which will be used to fetch the configuration
             * @param {string} cacheKey Key that will be used to store cached items to the cache
             * @param {IPnPClientStore} cacheStore OPTIONAL storage, which will be used to store cached settings.
             */
            function CachingConfigurationProvider(wrappedProvider, cacheKey, cacheStore) {
                this.wrappedProvider = wrappedProvider;
                this.store = (cacheStore) ? cacheStore : this.selectPnPCache();
                this.cacheKey = "_configcache_" + cacheKey;
            }
            /**
             * Gets the wrapped configuration providers
             *
             * @return {IConfigurationProvider} Wrapped configuration provider
             */
            CachingConfigurationProvider.prototype.getWrappedProvider = function () {
                return this.wrappedProvider;
            };
            /**
             * Loads the configuration values either from the cache or from the wrapped provider
             *
             * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
             */
            CachingConfigurationProvider.prototype.getConfiguration = function () {
                var _this = this;
                // Cache not available, pass control to  the wrapped provider
                if ((!this.store) || (!this.store.enabled)) {
                    return this.wrappedProvider.getConfiguration();
                }
                // Value is found in cache, return it directly
                var cachedConfig = this.store.get(this.cacheKey);
                if (cachedConfig) {
                    return new Promise(function (resolve) {
                        resolve(cachedConfig);
                    });
                }
                // Get and cache value from the wrapped provider
                var providerPromise = this.wrappedProvider.getConfiguration();
                providerPromise.then(function (providedConfig) {
                    _this.store.put(_this.cacheKey, providedConfig);
                });
                return providerPromise;
            };
            CachingConfigurationProvider.prototype.selectPnPCache = function () {
                var pnpCache = new storage.PnPClientStorage();
                if ((pnpCache.local) && (pnpCache.local.enabled)) {
                    return pnpCache.local;
                }
                if ((pnpCache.session) && (pnpCache.session.enabled)) {
                    return pnpCache.session;
                }
                throw new exceptions_1.NoCacheAvailableException();
            };
            return CachingConfigurationProvider;
        }());
        exports.default = CachingConfigurationProvider;


        /***/
    }),
    /* 18 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";
        /* WEBPACK VAR INJECTION */(function (global) {
            Object.defineProperty(exports, "__esModule", { value: true });
            /**
             * Makes requests using the fetch API
             */
            var FetchClient = (function () {
                function FetchClient() {
                }
                FetchClient.prototype.fetch = function (url, options) {
                    return global.fetch(url, options);
                };
                return FetchClient;
            }());
            exports.FetchClient = FetchClient;

            /* WEBPACK VAR INJECTION */
        }.call(exports, __webpack_require__(30)))

        /***/
    }),
    /* 19 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var storage_1 = __webpack_require__(9);
        var util_1 = __webpack_require__(0);
        var pnplibconfig_1 = __webpack_require__(4);
        var CachingOptions = (function () {
            function CachingOptions(key) {
                this.key = key;
                this.expiration = util_1.Util.dateAdd(new Date(), "second", pnplibconfig_1.RuntimeConfig.defaultCachingTimeoutSeconds);
                this.storeName = pnplibconfig_1.RuntimeConfig.defaultCachingStore;
            }
            Object.defineProperty(CachingOptions.prototype, "store", {
                get: function () {
                    if (this.storeName === "local") {
                        return CachingOptions.storage.local;
                    }
                    else {
                        return CachingOptions.storage.session;
                    }
                },
                enumerable: true,
                configurable: true
            });
            return CachingOptions;
        }());
        CachingOptions.storage = new storage_1.PnPClientStorage();
        exports.CachingOptions = CachingOptions;
        var CachingParserWrapper = (function () {
            function CachingParserWrapper(_parser, _cacheOptions) {
                this._parser = _parser;
                this._cacheOptions = _cacheOptions;
            }
            CachingParserWrapper.prototype.parse = function (response) {
                var _this = this;
                // add this to the cache based on the options
                return this._parser.parse(response).then(function (data) {
                    if (_this._cacheOptions.store !== null) {
                        _this._cacheOptions.store.put(_this._cacheOptions.key, data, _this._cacheOptions.expiration);
                    }
                    return data;
                });
            };
            return CachingParserWrapper;
        }());
        exports.CachingParserWrapper = CachingParserWrapper;


        /***/
    }),
    /* 20 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        /**
         * Describes a collection of List objects
         *
         */
        var Features = (function (_super) {
            __extends(Features, _super);
            /**
             * Creates a new instance of the Lists class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Features(baseUrl, path) {
                if (path === void 0) { path = "features"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a list from the collection by guid id
             *
             * @param id The Id of the feature (GUID)
             */
            Features.prototype.getById = function (id) {
                var feature = new Feature(this);
                feature.concat("('" + id + "')");
                return feature;
            };
            /**
             * Adds a new list to the collection
             *
             * @param id The Id of the feature (GUID)
             * @param force If true the feature activation will be forced
             */
            Features.prototype.add = function (id, force) {
                var _this = this;
                if (force === void 0) { force = false; }
                var adder = new Features(this, "add");
                return adder.post({
                    body: JSON.stringify({
                        featdefScope: 0,
                        featureId: id,
                        force: force,
                    }),
                }).then(function (data) {
                    return {
                        data: data,
                        feature: _this.getById(id),
                    };
                });
            };
            /**
             * Removes (deactivates) a feature from the collection
             *
             * @param id The Id of the feature (GUID)
             * @param force If true the feature deactivation will be forced
             */
            Features.prototype.remove = function (id, force) {
                if (force === void 0) { force = false; }
                var remover = new Features(this, "remove");
                return remover.post({
                    body: JSON.stringify({
                        featureId: id,
                        force: force,
                    }),
                });
            };
            return Features;
        }(queryable_1.QueryableCollection));
        exports.Features = Features;
        var Feature = (function (_super) {
            __extends(Feature, _super);
            /**
             * Creates a new instance of the Lists class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Feature(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Removes (deactivates) a feature from the collection
             *
             * @param force If true the feature deactivation will be forced
             */
            Feature.prototype.deactivate = function (force) {
                var _this = this;
                if (force === void 0) { force = false; }
                var removeDependency = this.addBatchDependency();
                var idGet = new Feature(this).select("DefinitionId");
                return idGet.getAs().then(function (feature) {
                    var promise = _this.getParent(Features, _this.parentUrl, "").remove(feature.DefinitionId, force);
                    removeDependency();
                    return promise;
                });
            };
            return Feature;
        }(queryable_1.QueryableInstance));
        exports.Feature = Feature;


        /***/
    }),
    /* 21 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var util_1 = __webpack_require__(0);
        var Types = __webpack_require__(29);
        /**
         * Describes a collection of Field objects
         *
         */
        var Fields = (function (_super) {
            __extends(Fields, _super);
            /**
             * Creates a new instance of the Fields class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Fields(baseUrl, path) {
                if (path === void 0) { path = "fields"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a field from the collection by title
             *
             * @param title The case-sensitive title of the field
             */
            Fields.prototype.getByTitle = function (title) {
                return new Field(this, "getByTitle('" + title + "')");
            };
            /**
             * Gets a field from the collection by using internal name or title
             *
             * @param name The case-sensitive internal name or title of the field
             */
            Fields.prototype.getByInternalNameOrTitle = function (name) {
                return new Field(this, "getByInternalNameOrTitle('" + name + "')");
            };
            /**
             * Gets a list from the collection by guid id
             *
             * @param title The Id of the list
             */
            Fields.prototype.getById = function (id) {
                var f = new Field(this);
                f.concat("('" + id + "')");
                return f;
            };
            /**
             * Creates a field based on the specified schema
             */
            Fields.prototype.createFieldAsXml = function (xml) {
                var _this = this;
                var info;
                if (typeof xml === "string") {
                    info = { SchemaXml: xml };
                }
                else {
                    info = xml;
                }
                var postBody = JSON.stringify({
                    "parameters": util_1.Util.extend({
                        "__metadata": {
                            "type": "SP.XmlSchemaFieldCreationInformation",
                        },
                    }, info),
                });
                var q = new Fields(this, "createfieldasxml");
                return q.postAs({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        field: _this.getById(data.Id),
                    };
                });
            };
            /**
             * Adds a new list to the collection
             *
             * @param title The new field's title
             * @param fieldType The new field's type (ex: SP.FieldText)
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             */
            Fields.prototype.add = function (title, fieldType, properties) {
                var _this = this;
                if (properties === void 0) { properties = {}; }
                var postBody = JSON.stringify(util_1.Util.extend({
                    "Title": title,
                    "__metadata": { "type": fieldType },
                }, properties));
                return this.postAs({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        field: _this.getById(data.Id),
                    };
                });
            };
            /**
             * Adds a new SP.FieldText to the collection
             *
             * @param title The field title
             * @param maxLength The maximum number of characters allowed in the value of the field.
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             */
            Fields.prototype.addText = function (title, maxLength, properties) {
                if (maxLength === void 0) { maxLength = 255; }
                var props = {
                    FieldTypeKind: 2,
                    MaxLength: maxLength,
                };
                return this.add(title, "SP.FieldText", util_1.Util.extend(props, properties));
            };
            /**
             * Adds a new SP.FieldCalculated to the collection
             *
             * @param title The field title.
             * @param formula The formula for the field.
             * @param dateFormat The date and time format that is displayed in the field.
             * @param outputType Specifies the output format for the field. Represents a FieldType value.
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             */
            Fields.prototype.addCalculated = function (title, formula, dateFormat, outputType, properties) {
                if (outputType === void 0) { outputType = Types.FieldTypes.Text; }
                var props = {
                    DateFormat: dateFormat,
                    FieldTypeKind: 17,
                    Formula: formula,
                    OutputType: outputType,
                };
                return this.add(title, "SP.FieldCalculated", util_1.Util.extend(props, properties));
            };
            /**
             * Adds a new SP.FieldDateTime to the collection
             *
             * @param title The field title
             * @param displayFormat The format of the date and time that is displayed in the field.
             * @param calendarType Specifies the calendar type of the field.
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             */
            Fields.prototype.addDateTime = function (title, displayFormat, calendarType, friendlyDisplayFormat, properties) {
                if (displayFormat === void 0) { displayFormat = Types.DateTimeFieldFormatType.DateOnly; }
                if (calendarType === void 0) { calendarType = Types.CalendarType.Gregorian; }
                if (friendlyDisplayFormat === void 0) { friendlyDisplayFormat = 0; }
                var props = {
                    DateTimeCalendarType: calendarType,
                    DisplayFormat: displayFormat,
                    FieldTypeKind: 4,
                    FriendlyDisplayFormat: friendlyDisplayFormat,
                };
                return this.add(title, "SP.FieldDateTime", util_1.Util.extend(props, properties));
            };
            /**
             * Adds a new SP.FieldNumber to the collection
             *
             * @param title The field title
             * @param minValue The field's minimum value
             * @param maxValue The field's maximum value
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             */
            Fields.prototype.addNumber = function (title, minValue, maxValue, properties) {
                var props = { FieldTypeKind: 9 };
                if (typeof minValue !== "undefined") {
                    props = util_1.Util.extend({ MinimumValue: minValue }, props);
                }
                if (typeof maxValue !== "undefined") {
                    props = util_1.Util.extend({ MaximumValue: maxValue }, props);
                }
                return this.add(title, "SP.FieldNumber", util_1.Util.extend(props, properties));
            };
            /**
             * Adds a new SP.FieldCurrency to the collection
             *
             * @param title The field title
             * @param minValue The field's minimum value
             * @param maxValue The field's maximum value
             * @param currencyLocalId Specifies the language code identifier (LCID) used to format the value of the field
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             */
            Fields.prototype.addCurrency = function (title, minValue, maxValue, currencyLocalId, properties) {
                if (currencyLocalId === void 0) { currencyLocalId = 1033; }
                var props = {
                    CurrencyLocaleId: currencyLocalId,
                    FieldTypeKind: 10,
                };
                if (typeof minValue !== "undefined") {
                    props = util_1.Util.extend({ MinimumValue: minValue }, props);
                }
                if (typeof maxValue !== "undefined") {
                    props = util_1.Util.extend({ MaximumValue: maxValue }, props);
                }
                return this.add(title, "SP.FieldCurrency", util_1.Util.extend(props, properties));
            };
            /**
             * Adds a new SP.FieldMultiLineText to the collection
             *
             * @param title The field title
             * @param numberOfLines Specifies the number of lines of text to display for the field.
             * @param richText Specifies whether the field supports rich formatting.
             * @param restrictedMode Specifies whether the field supports a subset of rich formatting.
             * @param appendOnly Specifies whether all changes to the value of the field are displayed in list forms.
             * @param allowHyperlink Specifies whether a hyperlink is allowed as a value of the field.
             * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
             *
             */
            Fields.prototype.addMultilineText = function (title, numberOfLines, richText, restrictedMode, appendOnly, allowHyperlink, properties) {
                if (numberOfLines === void 0) { numberOfLines = 6; }
                if (richText === void 0) { richText = true; }
                if (restrictedMode === void 0) { restrictedMode = false; }
                if (appendOnly === void 0) { appendOnly = false; }
                if (allowHyperlink === void 0) { allowHyperlink = true; }
                var props = {
                    AllowHyperlink: allowHyperlink,
                    AppendOnly: appendOnly,
                    FieldTypeKind: 3,
                    NumberOfLines: numberOfLines,
                    RestrictedMode: restrictedMode,
                    RichText: richText,
                };
                return this.add(title, "SP.FieldMultiLineText", util_1.Util.extend(props, properties));
            };
            /**
             * Adds a new SP.FieldUrl to the collection
             *
             * @param title The field title
             */
            Fields.prototype.addUrl = function (title, displayFormat, properties) {
                if (displayFormat === void 0) { displayFormat = Types.UrlFieldFormatType.Hyperlink; }
                var props = {
                    DisplayFormat: displayFormat,
                    FieldTypeKind: 11,
                };
                return this.add(title, "SP.FieldUrl", util_1.Util.extend(props, properties));
            };
            return Fields;
        }(queryable_1.QueryableCollection));
        exports.Fields = Fields;
        /**
         * Describes a single of Field instance
         *
         */
        var Field = (function (_super) {
            __extends(Field, _super);
            /**
             * Creates a new instance of the Field class
             *
             * @param baseUrl The url or Queryable which forms the parent of this field instance
             */
            function Field(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Updates this field intance with the supplied properties
             *
             * @param properties A plain object hash of values to update for the list
             * @param fieldType The type value, required to update child field type properties
             */
            Field.prototype.update = function (properties, fieldType) {
                var _this = this;
                if (fieldType === void 0) { fieldType = "SP.Field"; }
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": fieldType },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    return {
                        data: data,
                        field: _this,
                    };
                });
            };
            /**
             * Delete this fields
             *
             */
            Field.prototype.delete = function () {
                return this.post({
                    headers: {
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            /**
             * Sets the value of the ShowInDisplayForm property for this field.
             */
            Field.prototype.setShowInDisplayForm = function (show) {
                var q = new Field(this, "setshowindisplayform(" + show + ")");
                return q.post();
            };
            /**
             * Sets the value of the ShowInEditForm property for this field.
             */
            Field.prototype.setShowInEditForm = function (show) {
                var q = new Field(this, "setshowineditform(" + show + ")");
                return q.post();
            };
            /**
             * Sets the value of the ShowInNewForm property for this field.
             */
            Field.prototype.setShowInNewForm = function (show) {
                var q = new Field(this, "setshowinnewform(" + show + ")");
                return q.post();
            };
            return Field;
        }(queryable_1.QueryableInstance));
        exports.Field = Field;


        /***/
    }),
    /* 22 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var files_1 = __webpack_require__(7);
        /**
         * Describes a collection of Folder objects
         *
         */
        var Folders = (function (_super) {
            __extends(Folders, _super);
            /**
             * Creates a new instance of the Folders class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Folders(baseUrl, path) {
                if (path === void 0) { path = "folders"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a folder by folder name
             *
             */
            Folders.prototype.getByName = function (name) {
                var f = new Folder(this);
                f.concat("('" + name + "')");
                return f;
            };
            /**
             * Adds a new folder to the current folder (relative) or any folder (absolute)
             *
             * @param url The relative or absolute url where the new folder will be created. Urls starting with a forward slash are absolute.
             * @returns The new Folder and the raw response.
             */
            Folders.prototype.add = function (url) {
                var _this = this;
                return new Folders(this, "add('" + url + "')").post().then(function (response) {
                    return {
                        data: response,
                        folder: _this.getByName(url),
                    };
                });
            };
            return Folders;
        }(queryable_1.QueryableCollection));
        exports.Folders = Folders;
        /**
         * Describes a single Folder instance
         *
         */
        var Folder = (function (_super) {
            __extends(Folder, _super);
            //
            // TODO:
            //      Properties (https://msdn.microsoft.com/en-us/library/office/dn450841.aspx#bk_FolderProperties)
            //          UniqueContentTypeOrder (setter)
            //          WelcomePage (setter)
            //
            /**
             * Creates a new instance of the Folder class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function Folder(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(Folder.prototype, "contentTypeOrder", {
                /**
                 * Specifies the sequence in which content types are displayed.
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "contentTypeOrder");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "files", {
                /**
                 * Gets this folder's files
                 *
                 */
                get: function () {
                    return new files_1.Files(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "folders", {
                /**
                 * Gets this folder's sub folders
                 *
                 */
                get: function () {
                    return new Folders(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "listItemAllFields", {
                /**
                 * Gets this folder's list item field values
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "listItemAllFields");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "parentFolder", {
                /**
                 * Gets the parent folder, if available
                 *
                 */
                get: function () {
                    return new Folder(this, "parentFolder");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "properties", {
                /**
                 * Gets this folder's properties
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableInstance(this, "properties");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "serverRelativeUrl", {
                /**
                 * Gets this folder's server relative url
                 *
                 */
                get: function () {
                    return new queryable_1.Queryable(this, "serverRelativeUrl");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Folder.prototype, "uniqueContentTypeOrder", {
                /**
                 * Gets a value that specifies the content type order.
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "uniqueContentTypeOrder");
                },
                enumerable: true,
                configurable: true
            });
            /**
            * Delete this folder
            *
            * @param eTag Value used in the IF-Match header, by default "*"
            */
            Folder.prototype.delete = function (eTag) {
                if (eTag === void 0) { eTag = "*"; }
                return new Folder(this).post({
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            /**
             * Moves the folder to the Recycle Bin and returns the identifier of the new Recycle Bin item.
             */
            Folder.prototype.recycle = function () {
                return new Folder(this, "recycle").post();
            };
            return Folder;
        }(queryable_1.QueryableInstance));
        exports.Folder = Folder;


        /***/
    }),
    /* 23 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var queryablesecurable_1 = __webpack_require__(12);
        var folders_1 = __webpack_require__(22);
        var files_1 = __webpack_require__(7);
        var contenttypes_1 = __webpack_require__(11);
        var util_1 = __webpack_require__(0);
        var odata_1 = __webpack_require__(3);
        var attachmentfiles_1 = __webpack_require__(39);
        var lists_1 = __webpack_require__(8);
        /**
         * Describes a collection of Item objects
         *
         */
        var Items = (function (_super) {
            __extends(Items, _super);
            /**
             * Creates a new instance of the Items class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Items(baseUrl, path) {
                if (path === void 0) { path = "items"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets an Item by id
             *
             * @param id The integer id of the item to retrieve
             */
            Items.prototype.getById = function (id) {
                var i = new Item(this);
                i.concat("(" + id + ")");
                return i;
            };
            /**
             * Skips the specified number of items (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#sectionSection6)
             *
             * @param skip The starting id where the page should start, use with top to specify pages
             */
            Items.prototype.skip = function (skip) {
                this._query.add("$skiptoken", encodeURIComponent("Paged=TRUE&p_ID=" + skip));
                return this;
            };
            /**
             * Gets a collection designed to aid in paging through data
             *
             */
            Items.prototype.getPaged = function () {
                return this.getAs(new PagedItemCollectionParser());
            };
            //
            /**
             * Adds a new item to the collection
             *
             * @param properties The new items's properties
             */
            Items.prototype.add = function (properties, listItemEntityTypeFullName) {
                var _this = this;
                if (properties === void 0) { properties = {}; }
                if (listItemEntityTypeFullName === void 0) { listItemEntityTypeFullName = null; }
                var removeDependency = this.addBatchDependency();
                return this.ensureListItemEntityTypeName(listItemEntityTypeFullName).then(function (listItemEntityType) {
                    var postBody = JSON.stringify(util_1.Util.extend({
                        "__metadata": { "type": listItemEntityType },
                    }, properties));
                    var promise = _this.postAs({ body: postBody }).then(function (data) {
                        return {
                            data: data,
                            item: _this.getById(data.Id),
                        };
                    });
                    removeDependency();
                    return promise;
                });
            };
            /**
             * Ensures we have the proper list item entity type name, either from the value provided or from the list
             *
             * @param candidatelistItemEntityTypeFullName The potential type name
             */
            Items.prototype.ensureListItemEntityTypeName = function (candidatelistItemEntityTypeFullName) {
                if (candidatelistItemEntityTypeFullName) {
                    return Promise.resolve(candidatelistItemEntityTypeFullName);
                }
                return this.getParent(lists_1.List).getListItemEntityTypeFullName();
            };
            return Items;
        }(queryable_1.QueryableCollection));
        exports.Items = Items;
        /**
         * Descrines a single Item instance
         *
         */
        var Item = (function (_super) {
            __extends(Item, _super);
            /**
             * Creates a new instance of the Items class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Item(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(Item.prototype, "attachmentFiles", {
                /**
                 * Gets the set of attachments for this item
                 *
                 */
                get: function () {
                    return new attachmentfiles_1.AttachmentFiles(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "contentType", {
                /**
                 * Gets the content type for this item
                 *
                 */
                get: function () {
                    return new contenttypes_1.ContentType(this, "ContentType");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "effectiveBasePermissions", {
                /**
                 * Gets the effective base permissions for the item
                 *
                 */
                get: function () {
                    return new queryable_1.Queryable(this, "EffectiveBasePermissions");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "effectiveBasePermissionsForUI", {
                /**
                 * Gets the effective base permissions for the item in a UI context
                 *
                 */
                get: function () {
                    return new queryable_1.Queryable(this, "EffectiveBasePermissionsForUI");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "fieldValuesAsHTML", {
                /**
                 * Gets the field values for this list item in their HTML representation
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableInstance(this, "FieldValuesAsHTML");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "fieldValuesAsText", {
                /**
                 * Gets the field values for this list item in their text representation
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableInstance(this, "FieldValuesAsText");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "fieldValuesForEdit", {
                /**
                 * Gets the field values for this list item for use in editing controls
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableInstance(this, "FieldValuesForEdit");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "folder", {
                /**
                 * Gets the folder associated with this list item (if this item represents a folder)
                 *
                 */
                get: function () {
                    return new folders_1.Folder(this, "folder");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Item.prototype, "file", {
                /**
                 * Gets the folder associated with this list item (if this item represents a folder)
                 *
                 */
                get: function () {
                    return new files_1.File(this, "file");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Updates this list intance with the supplied properties
             *
             * @param properties A plain object hash of values to update for the list
             * @param eTag Value used in the IF-Match header, by default "*"
             */
            Item.prototype.update = function (properties, eTag) {
                var _this = this;
                if (eTag === void 0) { eTag = "*"; }
                return new Promise(function (resolve, reject) {
                    var removeDependency = _this.addBatchDependency();
                    var parentList = _this.getParent(queryable_1.QueryableInstance, _this.parentUrl.substr(0, _this.parentUrl.lastIndexOf("/")));
                    parentList.select("ListItemEntityTypeFullName").getAs().then(function (d) {
                        var postBody = JSON.stringify(util_1.Util.extend({
                            "__metadata": { "type": d.ListItemEntityTypeFullName },
                        }, properties));
                        removeDependency();
                        return _this.post({
                            body: postBody,
                            headers: {
                                "IF-Match": eTag,
                                "X-HTTP-Method": "MERGE",
                            },
                        }, new ItemUpdatedParser()).then(function (data) {
                            resolve({
                                data: data,
                                item: _this,
                            });
                        });
                    }).catch(function (e) { return reject(e); });
                });
            };
            /**
             * Delete this item
             *
             * @param eTag Value used in the IF-Match header, by default "*"
             */
            Item.prototype.delete = function (eTag) {
                if (eTag === void 0) { eTag = "*"; }
                return this.post({
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            /**
             * Moves the list item to the Recycle Bin and returns the identifier of the new Recycle Bin item.
             */
            Item.prototype.recycle = function () {
                var i = new Item(this, "recycle");
                return i.post();
            };
            /**
             * Gets a string representation of the full URL to the WOPI frame.
             * If there is no associated WOPI application, or no associated action, an empty string is returned.
             *
             * @param action Display mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
             */
            Item.prototype.getWopiFrameUrl = function (action) {
                if (action === void 0) { action = 0; }
                var i = new Item(this, "getWOPIFrameUrl(@action)");
                i._query.add("@action", action);
                return i.post().then(function (data) {
                    return data.GetWOPIFrameUrl;
                });
            };
            /**
             * Validates and sets the values of the specified collection of fields for the list item.
             *
             * @param formValues The fields to change and their new values.
             * @param newDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
             */
            /* tslint:disable max-line-length */
            Item.prototype.validateUpdateListItem = function (formValues, newDocumentUpdate) {
                if (newDocumentUpdate === void 0) { newDocumentUpdate = false; }
                var postBody = JSON.stringify({ "formValues": formValues, bNewDocumentUpdate: newDocumentUpdate });
                var item = new Item(this, "validateupdatelistitem");
                return item.post({ body: postBody });
            };
            return Item;
        }(queryablesecurable_1.QueryableSecurable));
        exports.Item = Item;
        /**
         * Provides paging functionality for list items
         */
        var PagedItemCollection = (function () {
            function PagedItemCollection(nextUrl, results) {
                this.nextUrl = nextUrl;
                this.results = results;
            }
            Object.defineProperty(PagedItemCollection.prototype, "hasNext", {
                /**
                 * If true there are more results available in the set, otherwise there are not
                 */
                get: function () {
                    return typeof this.nextUrl === "string" && this.nextUrl.length > 0;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Gets the next set of results, or resolves to null if no results are available
             */
            PagedItemCollection.prototype.getNext = function () {
                if (this.hasNext) {
                    var items = new Items(this.nextUrl, null);
                    return items.getPaged();
                }
                return new Promise(function (r) { return r(null); });
            };
            return PagedItemCollection;
        }());
        exports.PagedItemCollection = PagedItemCollection;
        var PagedItemCollectionParser = (function (_super) {
            __extends(PagedItemCollectionParser, _super);
            function PagedItemCollectionParser() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            PagedItemCollectionParser.prototype.parse = function (r) {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    if (_this.handleError(r, reject)) {
                        r.json().then(function (json) {
                            var nextUrl = json.hasOwnProperty("d") && json.d.hasOwnProperty("__next") ? json.d.__next : json["odata.nextLink"];
                            resolve(new PagedItemCollection(nextUrl, _this.parseODataJSON(json)));
                        });
                    }
                });
            };
            return PagedItemCollectionParser;
        }(odata_1.ODataParserBase));
        var ItemUpdatedParser = (function (_super) {
            __extends(ItemUpdatedParser, _super);
            function ItemUpdatedParser() {
                return _super !== null && _super.apply(this, arguments) || this;
            }
            ItemUpdatedParser.prototype.parse = function (r) {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    if (_this.handleError(r, reject)) {
                        resolve({
                            "odata.etag": r.headers.get("etag"),
                        });
                    }
                });
            };
            return ItemUpdatedParser;
        }(odata_1.ODataParserBase));


        /***/
    }),
    /* 24 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        var queryable_1 = __webpack_require__(1);
        /**
         * Represents a collection of navigation nodes
         *
         */
        var NavigationNodes = (function (_super) {
            __extends(NavigationNodes, _super);
            function NavigationNodes(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a navigation node by id
             *
             * @param id The id of the node
             */
            NavigationNodes.prototype.getById = function (id) {
                var node = new NavigationNode(this);
                node.concat("(" + id + ")");
                return node;
            };
            /**
             * Adds a new node to the collection
             *
             * @param title Display name of the node
             * @param url The url of the node
             * @param visible If true the node is visible, otherwise it is hidden (default: true)
             */
            NavigationNodes.prototype.add = function (title, url, visible) {
                var _this = this;
                if (visible === void 0) { visible = true; }
                var postBody = JSON.stringify({
                    IsVisible: visible,
                    Title: title,
                    Url: url,
                    "__metadata": { "type": "SP.NavigationNode" },
                });
                var adder = new NavigationNodes(this);
                return adder.post({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        node: _this.getById(data.Id),
                    };
                });
            };
            /**
             * Moves a node to be after another node in the navigation
             *
             * @param nodeId Id of the node to move
             * @param previousNodeId Id of the node after which we move the node specified by nodeId
             */
            NavigationNodes.prototype.moveAfter = function (nodeId, previousNodeId) {
                var postBody = JSON.stringify({
                    nodeId: nodeId,
                    previousNodeId: previousNodeId,
                });
                var mover = new NavigationNodes(this, "MoveAfter");
                return mover.post({ body: postBody });
            };
            return NavigationNodes;
        }(queryable_1.QueryableCollection));
        exports.NavigationNodes = NavigationNodes;
        var NavigationNode = (function (_super) {
            __extends(NavigationNode, _super);
            function NavigationNode(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(NavigationNode.prototype, "children", {
                /**
                 * Represents the child nodes of this node
                 */
                get: function () {
                    return new NavigationNodes(this, "Children");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Updates this node based on the supplied properties
             *
             * @param properties The hash of key/value pairs to update
             */
            NavigationNode.prototype.update = function (properties) {
                var _this = this;
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": "SP.NavigationNode" },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    return {
                        data: data,
                        node: _this,
                    };
                });
            };
            /**
             * Deletes this node and any child nodes
             */
            NavigationNode.prototype.delete = function () {
                return _super.prototype.delete.call(this);
            };
            return NavigationNode;
        }(queryable_1.QueryableInstance));
        exports.NavigationNode = NavigationNode;
        /**
         * Exposes the navigation components
         *
         */
        var Navigation = (function (_super) {
            __extends(Navigation, _super);
            /**
             * Creates a new instance of the Lists class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Navigation(baseUrl, path) {
                if (path === void 0) { path = "navigation"; }
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(Navigation.prototype, "quicklaunch", {
                /**
                 * Gets the quicklaunch navigation for the current context
                 *
                 */
                get: function () {
                    return new NavigationNodes(this, "quicklaunch");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Navigation.prototype, "topNavigationBar", {
                /**
                 * Gets the top bar navigation navigation for the current context
                 *
                 */
                get: function () {
                    return new NavigationNodes(this, "topnavigationbar");
                },
                enumerable: true,
                configurable: true
            });
            return Navigation;
        }(queryable_1.Queryable));
        exports.Navigation = Navigation;


        /***/
    }),
    /* 25 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var util_1 = __webpack_require__(0);
        /**
         * Describes the search API
         *
         */
        var Search = (function (_super) {
            __extends(Search, _super);
            /**
             * Creates a new instance of the Search class
             *
             * @param baseUrl The url for the search context
             * @param query The SearchQuery object to execute
             */
            function Search(baseUrl, path) {
                if (path === void 0) { path = "_api/search/postquery"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * .......
             * @returns Promise
             */
            Search.prototype.execute = function (query) {
                var formattedBody;
                formattedBody = query;
                if (formattedBody.SelectProperties) {
                    formattedBody.SelectProperties = { results: query.SelectProperties };
                }
                if (formattedBody.RefinementFilters) {
                    formattedBody.RefinementFilters = { results: query.RefinementFilters };
                }
                if (formattedBody.SortList) {
                    formattedBody.SortList = { results: query.SortList };
                }
                if (formattedBody.HithighlightedProperties) {
                    formattedBody.HithighlightedProperties = { results: query.HithighlightedProperties };
                }
                if (formattedBody.ReorderingRules) {
                    formattedBody.ReorderingRules = { results: query.ReorderingRules };
                }
                if (formattedBody.Properties) {
                    formattedBody.Properties = { results: query.Properties };
                }
                var postBody = JSON.stringify({
                    request: util_1.Util.extend({
                        "__metadata": { "type": "Microsoft.Office.Server.Search.REST.SearchRequest" },
                    }, formattedBody),
                });
                return this.post({ body: postBody }).then(function (data) { return new SearchResults(data); });
            };
            return Search;
        }(queryable_1.QueryableInstance));
        exports.Search = Search;
        /**
         * Describes the SearchResults class, which returns the formatted and raw version of the query response
         */
        var SearchResults = (function () {
            /**
             * Creates a new instance of the SearchResult class
             *
             */
            function SearchResults(rawResponse) {
                var response = rawResponse.postquery ? rawResponse.postquery : rawResponse;
                this.PrimarySearchResults = this.formatSearchResults(response.PrimaryQueryResult.RelevantResults.Table.Rows);
                this.RawSearchResults = response;
                this.ElapsedTime = response.ElapsedTime;
                this.RowCount = response.PrimaryQueryResult.RelevantResults.RowCount;
                this.TotalRows = response.PrimaryQueryResult.RelevantResults.TotalRows;
                this.TotalRowsIncludingDuplicates = response.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates;
            }
            /**
             * Formats a search results array
             *
             * @param rawResults The array to process
             */
            SearchResults.prototype.formatSearchResults = function (rawResults) {
                var results = new Array(), tempResults = rawResults.results ? rawResults.results : rawResults;
                for (var _i = 0, tempResults_1 = tempResults; _i < tempResults_1.length; _i++) {
                    var i = tempResults_1[_i];
                    results.push(new SearchResult(i.Cells));
                }
                return results;
            };
            return SearchResults;
        }());
        exports.SearchResults = SearchResults;
        /**
         * Describes the SearchResult class
         */
        var SearchResult = (function () {
            /**
             * Creates a new instance of the SearchResult class
             *
             */
            function SearchResult(rawItem) {
                var item = rawItem.results ? rawItem.results : rawItem;
                for (var _i = 0, item_1 = item; _i < item_1.length; _i++) {
                    var i = item_1[_i];
                    Object.defineProperty(this, i.Key, {
                        configurable: false,
                        enumerable: false,
                        value: i.Value,
                        writable: false,
                    });
                }
            }
            return SearchResult;
        }());
        exports.SearchResult = SearchResult;
        /**
         * defines the SortDirection enum
         */
        var SortDirection;
        (function (SortDirection) {
            SortDirection[SortDirection["Ascending"] = 0] = "Ascending";
            SortDirection[SortDirection["Descending"] = 1] = "Descending";
            SortDirection[SortDirection["FQLFormula"] = 2] = "FQLFormula";
        })(SortDirection = exports.SortDirection || (exports.SortDirection = {}));
        /**
         * defines the ReorderingRuleMatchType  enum
         */
        var ReorderingRuleMatchType;
        (function (ReorderingRuleMatchType) {
            ReorderingRuleMatchType[ReorderingRuleMatchType["ResultContainsKeyword"] = 0] = "ResultContainsKeyword";
            ReorderingRuleMatchType[ReorderingRuleMatchType["TitleContainsKeyword"] = 1] = "TitleContainsKeyword";
            ReorderingRuleMatchType[ReorderingRuleMatchType["TitleMatchesKeyword"] = 2] = "TitleMatchesKeyword";
            ReorderingRuleMatchType[ReorderingRuleMatchType["UrlStartsWith"] = 3] = "UrlStartsWith";
            ReorderingRuleMatchType[ReorderingRuleMatchType["UrlExactlyMatches"] = 4] = "UrlExactlyMatches";
            ReorderingRuleMatchType[ReorderingRuleMatchType["ContentTypeIs"] = 5] = "ContentTypeIs";
            ReorderingRuleMatchType[ReorderingRuleMatchType["FileExtensionMatches"] = 6] = "FileExtensionMatches";
            ReorderingRuleMatchType[ReorderingRuleMatchType["ResultHasTag"] = 7] = "ResultHasTag";
            ReorderingRuleMatchType[ReorderingRuleMatchType["ManualCondition"] = 8] = "ManualCondition";
        })(ReorderingRuleMatchType = exports.ReorderingRuleMatchType || (exports.ReorderingRuleMatchType = {}));
        /**
         * Specifies the type value for the property
         */
        var QueryPropertyValueType;
        (function (QueryPropertyValueType) {
            QueryPropertyValueType[QueryPropertyValueType["None"] = 0] = "None";
            QueryPropertyValueType[QueryPropertyValueType["StringType"] = 1] = "StringType";
            QueryPropertyValueType[QueryPropertyValueType["Int32TYpe"] = 2] = "Int32TYpe";
            QueryPropertyValueType[QueryPropertyValueType["BooleanType"] = 3] = "BooleanType";
            QueryPropertyValueType[QueryPropertyValueType["StringArrayType"] = 4] = "StringArrayType";
            QueryPropertyValueType[QueryPropertyValueType["UnSupportedType"] = 5] = "UnSupportedType";
        })(QueryPropertyValueType = exports.QueryPropertyValueType || (exports.QueryPropertyValueType = {}));


        /***/
    }),
    /* 26 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var SearchSuggest = (function (_super) {
            __extends(SearchSuggest, _super);
            function SearchSuggest(baseUrl, path) {
                if (path === void 0) { path = "_api/search/suggest"; }
                return _super.call(this, baseUrl, path) || this;
            }
            SearchSuggest.prototype.execute = function (query) {
                this.mapQueryToQueryString(query);
                return this.get().then(function (response) { return new SearchSuggestResult(response); });
            };
            SearchSuggest.prototype.mapQueryToQueryString = function (query) {
                this.query.add("querytext", "'" + query.querytext + "'");
                if (query.hasOwnProperty("count")) {
                    this.query.add("inumberofquerysuggestions", query.count.toString());
                }
                if (query.hasOwnProperty("personalCount")) {
                    this.query.add("inumberofresultsuggestions", query.personalCount.toString());
                }
                if (query.hasOwnProperty("preQuery")) {
                    this.query.add("fprequerysuggestions", query.preQuery.toString());
                }
                if (query.hasOwnProperty("hitHighlighting")) {
                    this.query.add("fhithighlighting", query.hitHighlighting.toString());
                }
                if (query.hasOwnProperty("capitalize")) {
                    this.query.add("fcapitalizefirstletters", query.capitalize.toString());
                }
                if (query.hasOwnProperty("culture")) {
                    this.query.add("culture", query.culture.toString());
                }
                if (query.hasOwnProperty("stemming")) {
                    this.query.add("enablestemming", query.stemming.toString());
                }
                if (query.hasOwnProperty("includePeople")) {
                    this.query.add("showpeoplenamesuggestions", query.includePeople.toString());
                }
                if (query.hasOwnProperty("queryRules")) {
                    this.query.add("enablequeryrules", query.queryRules.toString());
                }
                if (query.hasOwnProperty("prefixMatch")) {
                    this.query.add("fprefixmatchallterms", query.prefixMatch.toString());
                }
            };
            return SearchSuggest;
        }(queryable_1.QueryableInstance));
        exports.SearchSuggest = SearchSuggest;
        var SearchSuggestResult = (function () {
            function SearchSuggestResult(json) {
                if (json.hasOwnProperty("suggest")) {
                    // verbose
                    this.PeopleNames = json.suggest.PeopleNames.results;
                    this.PersonalResults = json.suggest.PersonalResults.results;
                    this.Queries = json.suggest.Queries.results;
                }
                else {
                    this.PeopleNames = json.PeopleNames;
                    this.PersonalResults = json.PersonalResults;
                    this.Queries = json.Queries;
                }
            }
            return SearchSuggestResult;
        }());
        exports.SearchSuggestResult = SearchSuggestResult;


        /***/
    }),
    /* 27 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var webs_1 = __webpack_require__(16);
        var usercustomactions_1 = __webpack_require__(15);
        var odata_1 = __webpack_require__(3);
        var features_1 = __webpack_require__(20);
        /**
         * Describes a site collection
         *
         */
        var Site = (function (_super) {
            __extends(Site, _super);
            /**
             * Creates a new instance of the RoleAssignments class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Site(baseUrl, path) {
                if (path === void 0) { path = "_api/site"; }
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(Site.prototype, "rootWeb", {
                /**
                 * Gets the root web of the site collection
                 *
                 */
                get: function () {
                    return new webs_1.Web(this, "rootweb");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Site.prototype, "features", {
                /**
                 * Gets the active features for this site
                 *
                 */
                get: function () {
                    return new features_1.Features(this);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Site.prototype, "userCustomActions", {
                /**
                 * Get all custom actions on a site collection
                 *
                 */
                get: function () {
                    return new usercustomactions_1.UserCustomActions(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Gets the context information for the site.
             */
            Site.prototype.getContextInfo = function () {
                var q = new Site(this.parentUrl, "_api/contextinfo");
                return q.post().then(function (data) {
                    if (data.hasOwnProperty("GetContextWebInformation")) {
                        var info = data.GetContextWebInformation;
                        info.SupportedSchemaVersions = info.SupportedSchemaVersions.results;
                        return info;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Gets the document libraries on a site. Static method. (SharePoint Online only)
             *
             * @param absoluteWebUrl The absolute url of the web whose document libraries should be returned
             */
            Site.prototype.getDocumentLibraries = function (absoluteWebUrl) {
                var q = new queryable_1.Queryable("", "_api/sp.web.getdocumentlibraries(@v)");
                q.query.add("@v", "'" + absoluteWebUrl + "'");
                return q.get().then(function (data) {
                    if (data.hasOwnProperty("GetDocumentLibraries")) {
                        return data.GetDocumentLibraries;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Gets the site URL from a page URL.
             *
             * @param absolutePageUrl The absolute url of the page
             */
            Site.prototype.getWebUrlFromPageUrl = function (absolutePageUrl) {
                var q = new queryable_1.Queryable("", "_api/sp.web.getweburlfrompageurl(@v)");
                q.query.add("@v", "'" + absolutePageUrl + "'");
                return q.get().then(function (data) {
                    if (data.hasOwnProperty("GetWebUrlFromPageUrl")) {
                        return data.GetWebUrlFromPageUrl;
                    }
                    else {
                        return data;
                    }
                });
            };
            /**
             * Creates a new batch for requests within the context of context this site
             *
             */
            Site.prototype.createBatch = function () {
                return new odata_1.ODataBatch(this.parentUrl);
            };
            return Site;
        }(queryable_1.QueryableInstance));
        exports.Site = Site;


        /***/
    }),
    /* 28 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var sitegroups_1 = __webpack_require__(14);
        var util_1 = __webpack_require__(0);
        /**
         * Describes a collection of all site collection users
         *
         */
        var SiteUsers = (function (_super) {
            __extends(SiteUsers, _super);
            /**
             * Creates a new instance of the Users class
             *
             * @param baseUrl The url or Queryable which forms the parent of this user collection
             */
            function SiteUsers(baseUrl, path) {
                if (path === void 0) { path = "siteusers"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a user from the collection by email
             *
             * @param email The email of the user
             */
            SiteUsers.prototype.getByEmail = function (email) {
                return new SiteUser(this, "getByEmail('" + email + "')");
            };
            /**
             * Gets a user from the collection by id
             *
             * @param id The id of the user
             */
            SiteUsers.prototype.getById = function (id) {
                return new SiteUser(this, "getById(" + id + ")");
            };
            /**
             * Gets a user from the collection by login name
             *
             * @param loginName The email address of the user
             */
            SiteUsers.prototype.getByLoginName = function (loginName) {
                var su = new SiteUser(this);
                su.concat("(@v)");
                su.query.add("@v", encodeURIComponent(loginName));
                return su;
            };
            /**
             * Removes a user from the collection by id
             *
             * @param id The id of the user
             */
            SiteUsers.prototype.removeById = function (id) {
                var o = new SiteUsers(this, "removeById(" + id + ")");
                return o.post();
            };
            /**
             * Removes a user from the collection by login name
             *
             * @param loginName The login name of the user
             */
            SiteUsers.prototype.removeByLoginName = function (loginName) {
                var o = new SiteUsers(this, "removeByLoginName(@v)");
                o.query.add("@v", encodeURIComponent(loginName));
                return o.post();
            };
            /**
             * Add a user to a group
             *
             * @param loginName The login name of the user to add to the group
             *
             */
            SiteUsers.prototype.add = function (loginName) {
                var _this = this;
                var postBody = JSON.stringify({ "__metadata": { "type": "SP.User" }, LoginName: loginName });
                return this.post({ body: postBody }).then(function () { return _this.getByLoginName(loginName); });
            };
            return SiteUsers;
        }(queryable_1.QueryableCollection));
        exports.SiteUsers = SiteUsers;
        /**
         * Describes a single user
         *
         */
        var SiteUser = (function (_super) {
            __extends(SiteUser, _super);
            /**
             * Creates a new instance of the User class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, passes the path to the user
             */
            function SiteUser(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(SiteUser.prototype, "groups", {
                /**
                 * Get's the groups for this user.
                 *
                 */
                get: function () {
                    return new sitegroups_1.SiteGroups(this, "groups");
                },
                enumerable: true,
                configurable: true
            });
            /**
            * Updates this user instance with the supplied properties
            *
            * @param properties A plain object of property names and values to update for the user
            */
            SiteUser.prototype.update = function (properties) {
                var _this = this;
                var postBody = util_1.Util.extend({ "__metadata": { "type": "SP.User" } }, properties);
                return this.post({
                    body: JSON.stringify(postBody),
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    return {
                        data: data,
                        user: _this,
                    };
                });
            };
            /**
             * Delete this user
             *
             */
            SiteUser.prototype.delete = function () {
                return this.post({
                    headers: {
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            return SiteUser;
        }(queryable_1.QueryableInstance));
        exports.SiteUser = SiteUser;
        /**
         * Represents the current user
         */
        var CurrentUser = (function (_super) {
            __extends(CurrentUser, _super);
            function CurrentUser(baseUrl, path) {
                if (path === void 0) { path = "currentuser"; }
                return _super.call(this, baseUrl, path) || this;
            }
            return CurrentUser;
        }(queryable_1.QueryableInstance));
        exports.CurrentUser = CurrentUser;


        /***/
    }),
    /* 29 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";
        // reference: https://msdn.microsoft.com/en-us/library/office/dn600183.aspx

        Object.defineProperty(exports, "__esModule", { value: true });
        /**
         * Determines the display mode of the given control or view
         */
        var ControlMode;
        (function (ControlMode) {
            ControlMode[ControlMode["Display"] = 1] = "Display";
            ControlMode[ControlMode["Edit"] = 2] = "Edit";
            ControlMode[ControlMode["New"] = 3] = "New";
        })(ControlMode = exports.ControlMode || (exports.ControlMode = {}));
        /**
         * Specifies the type of the field.
         */
        var FieldTypes;
        (function (FieldTypes) {
            FieldTypes[FieldTypes["Invalid"] = 0] = "Invalid";
            FieldTypes[FieldTypes["Integer"] = 1] = "Integer";
            FieldTypes[FieldTypes["Text"] = 2] = "Text";
            FieldTypes[FieldTypes["Note"] = 3] = "Note";
            FieldTypes[FieldTypes["DateTime"] = 4] = "DateTime";
            FieldTypes[FieldTypes["Counter"] = 5] = "Counter";
            FieldTypes[FieldTypes["Choice"] = 6] = "Choice";
            FieldTypes[FieldTypes["Lookup"] = 7] = "Lookup";
            FieldTypes[FieldTypes["Boolean"] = 8] = "Boolean";
            FieldTypes[FieldTypes["Number"] = 9] = "Number";
            FieldTypes[FieldTypes["Currency"] = 10] = "Currency";
            FieldTypes[FieldTypes["URL"] = 11] = "URL";
            FieldTypes[FieldTypes["Computed"] = 12] = "Computed";
            FieldTypes[FieldTypes["Threading"] = 13] = "Threading";
            FieldTypes[FieldTypes["Guid"] = 14] = "Guid";
            FieldTypes[FieldTypes["MultiChoice"] = 15] = "MultiChoice";
            FieldTypes[FieldTypes["GridChoice"] = 16] = "GridChoice";
            FieldTypes[FieldTypes["Calculated"] = 17] = "Calculated";
            FieldTypes[FieldTypes["File"] = 18] = "File";
            FieldTypes[FieldTypes["Attachments"] = 19] = "Attachments";
            FieldTypes[FieldTypes["User"] = 20] = "User";
            FieldTypes[FieldTypes["Recurrence"] = 21] = "Recurrence";
            FieldTypes[FieldTypes["CrossProjectLink"] = 22] = "CrossProjectLink";
            FieldTypes[FieldTypes["ModStat"] = 23] = "ModStat";
            FieldTypes[FieldTypes["Error"] = 24] = "Error";
            FieldTypes[FieldTypes["ContentTypeId"] = 25] = "ContentTypeId";
            FieldTypes[FieldTypes["PageSeparator"] = 26] = "PageSeparator";
            FieldTypes[FieldTypes["ThreadIndex"] = 27] = "ThreadIndex";
            FieldTypes[FieldTypes["WorkflowStatus"] = 28] = "WorkflowStatus";
            FieldTypes[FieldTypes["AllDayEvent"] = 29] = "AllDayEvent";
            FieldTypes[FieldTypes["WorkflowEventType"] = 30] = "WorkflowEventType";
        })(FieldTypes = exports.FieldTypes || (exports.FieldTypes = {}));
        var DateTimeFieldFormatType;
        (function (DateTimeFieldFormatType) {
            DateTimeFieldFormatType[DateTimeFieldFormatType["DateOnly"] = 0] = "DateOnly";
            DateTimeFieldFormatType[DateTimeFieldFormatType["DateTime"] = 1] = "DateTime";
        })(DateTimeFieldFormatType = exports.DateTimeFieldFormatType || (exports.DateTimeFieldFormatType = {}));
        /**
         * Specifies the control settings while adding a field.
         */
        var AddFieldOptions;
        (function (AddFieldOptions) {
            /**
             *  Specify that a new field added to the list must also be added to the default content type in the site collection
             */
            AddFieldOptions[AddFieldOptions["DefaultValue"] = 0] = "DefaultValue";
            /**
             * Specify that a new field added to the list must also be added to the default content type in the site collection.
             */
            AddFieldOptions[AddFieldOptions["AddToDefaultContentType"] = 1] = "AddToDefaultContentType";
            /**
             * Specify that a new field must not be added to any other content type
             */
            AddFieldOptions[AddFieldOptions["AddToNoContentType"] = 2] = "AddToNoContentType";
            /**
             *  Specify that a new field that is added to the specified list must also be added to all content types in the site collection
             */
            AddFieldOptions[AddFieldOptions["AddToAllContentTypes"] = 4] = "AddToAllContentTypes";
            /**
             * Specify adding an internal field name hint for the purpose of avoiding possible database locking or field renaming operations
             */
            AddFieldOptions[AddFieldOptions["AddFieldInternalNameHint"] = 8] = "AddFieldInternalNameHint";
            /**
             * Specify that a new field that is added to the specified list must also be added to the default list view
             */
            AddFieldOptions[AddFieldOptions["AddFieldToDefaultView"] = 16] = "AddFieldToDefaultView";
            /**
             * Specify to confirm that no other field has the same display name
             */
            AddFieldOptions[AddFieldOptions["AddFieldCheckDisplayName"] = 32] = "AddFieldCheckDisplayName";
        })(AddFieldOptions = exports.AddFieldOptions || (exports.AddFieldOptions = {}));
        var CalendarType;
        (function (CalendarType) {
            CalendarType[CalendarType["Gregorian"] = 1] = "Gregorian";
            CalendarType[CalendarType["Japan"] = 3] = "Japan";
            CalendarType[CalendarType["Taiwan"] = 4] = "Taiwan";
            CalendarType[CalendarType["Korea"] = 5] = "Korea";
            CalendarType[CalendarType["Hijri"] = 6] = "Hijri";
            CalendarType[CalendarType["Thai"] = 7] = "Thai";
            CalendarType[CalendarType["Hebrew"] = 8] = "Hebrew";
            CalendarType[CalendarType["GregorianMEFrench"] = 9] = "GregorianMEFrench";
            CalendarType[CalendarType["GregorianArabic"] = 10] = "GregorianArabic";
            CalendarType[CalendarType["GregorianXLITEnglish"] = 11] = "GregorianXLITEnglish";
            CalendarType[CalendarType["GregorianXLITFrench"] = 12] = "GregorianXLITFrench";
            CalendarType[CalendarType["KoreaJapanLunar"] = 14] = "KoreaJapanLunar";
            CalendarType[CalendarType["ChineseLunar"] = 15] = "ChineseLunar";
            CalendarType[CalendarType["SakaEra"] = 16] = "SakaEra";
            CalendarType[CalendarType["UmAlQura"] = 23] = "UmAlQura";
        })(CalendarType = exports.CalendarType || (exports.CalendarType = {}));
        var UrlFieldFormatType;
        (function (UrlFieldFormatType) {
            UrlFieldFormatType[UrlFieldFormatType["Hyperlink"] = 0] = "Hyperlink";
            UrlFieldFormatType[UrlFieldFormatType["Image"] = 1] = "Image";
        })(UrlFieldFormatType = exports.UrlFieldFormatType || (exports.UrlFieldFormatType = {}));
        var PrincipalType;
        (function (PrincipalType) {
            PrincipalType[PrincipalType["None"] = 0] = "None";
            PrincipalType[PrincipalType["User"] = 1] = "User";
            PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
            PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
            PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
            PrincipalType[PrincipalType["All"] = 15] = "All";
        })(PrincipalType = exports.PrincipalType || (exports.PrincipalType = {}));
        var PageType;
        (function (PageType) {
            PageType[PageType["Invalid"] = -1] = "Invalid";
            PageType[PageType["DefaultView"] = 0] = "DefaultView";
            PageType[PageType["NormalView"] = 1] = "NormalView";
            PageType[PageType["DialogView"] = 2] = "DialogView";
            PageType[PageType["View"] = 3] = "View";
            PageType[PageType["DisplayForm"] = 4] = "DisplayForm";
            PageType[PageType["DisplayFormDialog"] = 5] = "DisplayFormDialog";
            PageType[PageType["EditForm"] = 6] = "EditForm";
            PageType[PageType["EditFormDialog"] = 7] = "EditFormDialog";
            PageType[PageType["NewForm"] = 8] = "NewForm";
            PageType[PageType["NewFormDialog"] = 9] = "NewFormDialog";
            PageType[PageType["SolutionForm"] = 10] = "SolutionForm";
            PageType[PageType["PAGE_MAXITEMS"] = 11] = "PAGE_MAXITEMS";
        })(PageType = exports.PageType || (exports.PageType = {}));


        /***/
    }),
    /* 30 */
    /***/ (function (module, exports) {

        var g;

        // This works in non-strict mode
        g = (function () {
            return this;
        })();

        try {
            // This works if eval is allowed (see CSP)
            g = g || Function("return this")() || (1, eval)("this");
        } catch (e) {
            // This works if the window reference is available
            if (typeof window === "object")
                g = window;
        }

        // g can still be undefined, but nothing to do about it...
        // We return undefined, instead of nothing here, so it's
        // easier to handle this case. if(!global) { ...}

        module.exports = g;


        /***/
    }),
    /* 31 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var collections_1 = __webpack_require__(6);
        /**
         * Class used to manage the current application settings
         *
         */
        var Settings = (function () {
            /**
             * Creates a new instance of the settings class
             *
             * @constructor
             */
            function Settings() {
                this._settings = new collections_1.Dictionary();
            }
            /**
             * Adds a new single setting, or overwrites a previous setting with the same key
             *
             * @param {string} key The key used to store this setting
             * @param {string} value The setting value to store
             */
            Settings.prototype.add = function (key, value) {
                this._settings.add(key, value);
            };
            /**
             * Adds a JSON value to the collection as a string, you must use getJSON to rehydrate the object when read
             *
             * @param {string} key The key used to store this setting
             * @param {any} value The setting value to store
             */
            Settings.prototype.addJSON = function (key, value) {
                this._settings.add(key, JSON.stringify(value));
            };
            /**
             * Applies the supplied hash to the setting collection overwriting any existing value, or created new values
             *
             * @param {TypedHash<any>} hash The set of values to add
             */
            Settings.prototype.apply = function (hash) {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    try {
                        _this._settings.merge(hash);
                        resolve();
                    }
                    catch (e) {
                        reject(e);
                    }
                });
            };
            /**
             * Loads configuration settings into the collection from the supplied provider and returns a Promise
             *
             * @param {IConfigurationProvider} provider The provider from which we will load the settings
             */
            Settings.prototype.load = function (provider) {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    provider.getConfiguration().then(function (value) {
                        _this._settings.merge(value);
                        resolve();
                    }).catch(function (reason) {
                        reject(reason);
                    });
                });
            };
            /**
             * Gets a value from the configuration
             *
             * @param {string} key The key whose value we want to return. Returns null if the key does not exist
             * @return {string} string value from the configuration
             */
            Settings.prototype.get = function (key) {
                return this._settings.get(key);
            };
            /**
             * Gets a JSON value, rehydrating the stored string to the original object
             *
             * @param {string} key The key whose value we want to return. Returns null if the key does not exist
             * @return {any} object from the configuration
             */
            Settings.prototype.getJSON = function (key) {
                var o = this.get(key);
                if (typeof o === "undefined" || o === null) {
                    return o;
                }
                return JSON.parse(o);
            };
            return Settings;
        }());
        exports.Settings = Settings;


        /***/
    }),
    /* 32 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var search_1 = __webpack_require__(25);
        var searchsuggest_1 = __webpack_require__(26);
        var site_1 = __webpack_require__(27);
        var webs_1 = __webpack_require__(16);
        var util_1 = __webpack_require__(0);
        var userprofiles_1 = __webpack_require__(44);
        var exceptions_1 = __webpack_require__(2);
        /**
         * Root of the SharePoint REST module
         */
        var Rest = (function () {
            function Rest() {
            }
            /**
             * Executes a search against this web context
             *
             * @param query The SearchQuery definition
             */
            Rest.prototype.searchSuggest = function (query) {
                var finalQuery;
                if (typeof query === "string") {
                    finalQuery = { querytext: query };
                }
                else {
                    finalQuery = query;
                }
                return new searchsuggest_1.SearchSuggest("").execute(finalQuery);
            };
            /**
             * Executes a search against this web context
             *
             * @param query The SearchQuery definition
             */
            Rest.prototype.search = function (query) {
                var finalQuery;
                if (typeof query === "string") {
                    finalQuery = { Querytext: query };
                }
                else {
                    finalQuery = query;
                }
                return new search_1.Search("").execute(finalQuery);
            };
            Object.defineProperty(Rest.prototype, "site", {
                /**
                 * Begins a site collection scoped REST request
                 *
                 */
                get: function () {
                    return new site_1.Site("");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Rest.prototype, "web", {
                /**
                 * Begins a web scoped REST request
                 *
                 */
                get: function () {
                    return new webs_1.Web("");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Rest.prototype, "profiles", {
                /**
                 * Access to user profile methods
                 *
                 */
                get: function () {
                    return new userprofiles_1.UserProfileQuery("");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Creates a new batch object for use with the Queryable.addToBatch method
             *
             */
            Rest.prototype.createBatch = function () {
                return this.web.createBatch();
            };
            /**
             * Begins a cross-domain, host site scoped REST request, for use in add-in webs
             *
             * @param addInWebUrl The absolute url of the add-in web
             * @param hostWebUrl The absolute url of the host web
             */
            Rest.prototype.crossDomainSite = function (addInWebUrl, hostWebUrl) {
                return this._cdImpl(site_1.Site, addInWebUrl, hostWebUrl, "site");
            };
            /**
             * Begins a cross-domain, host web scoped REST request, for use in add-in webs
             *
             * @param addInWebUrl The absolute url of the add-in web
             * @param hostWebUrl The absolute url of the host web
             */
            Rest.prototype.crossDomainWeb = function (addInWebUrl, hostWebUrl) {
                return this._cdImpl(webs_1.Web, addInWebUrl, hostWebUrl, "web");
            };
            /**
             * Implements the creation of cross domain REST urls
             *
             * @param factory The constructor of the object to create Site | Web
             * @param addInWebUrl The absolute url of the add-in web
             * @param hostWebUrl The absolute url of the host web
             * @param urlPart String part to append to the url "site" | "web"
             */
            Rest.prototype._cdImpl = function (factory, addInWebUrl, hostWebUrl, urlPart) {
                if (!util_1.Util.isUrlAbsolute(addInWebUrl)) {
                    throw new exceptions_1.UrlException("The addInWebUrl parameter must be an absolute url.");
                }
                if (!util_1.Util.isUrlAbsolute(hostWebUrl)) {
                    throw new exceptions_1.UrlException("The hostWebUrl parameter must be an absolute url.");
                }
                var url = util_1.Util.combinePaths(addInWebUrl, "_api/SP.AppContextSite(@target)");
                var instance = new factory(url, urlPart);
                instance.query.add("@target", "'" + encodeURIComponent(hostWebUrl) + "'");
                return instance;
            };
            return Rest;
        }());
        exports.Rest = Rest;


        /***/
    }),
    /* 33 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        function __export(m) {
            for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
        }
        Object.defineProperty(exports, "__esModule", { value: true });
        __export(__webpack_require__(41));
        var httpclient_1 = __webpack_require__(10);
        exports.HttpClient = httpclient_1.HttpClient;
        var sprequestexecutorclient_1 = __webpack_require__(38);
        exports.SPRequestExecutorClient = sprequestexecutorclient_1.SPRequestExecutorClient;
        var nodefetchclient_1 = __webpack_require__(37);
        exports.NodeFetchClient = nodefetchclient_1.NodeFetchClient;
        var fetchclient_1 = __webpack_require__(18);
        exports.FetchClient = fetchclient_1.FetchClient;
        __export(__webpack_require__(34));
        var collections_1 = __webpack_require__(6);
        exports.Dictionary = collections_1.Dictionary;
        var util_1 = __webpack_require__(0);
        exports.Util = util_1.Util;
        __export(__webpack_require__(5));
        __export(__webpack_require__(2));


        /***/
    }),
    /* 34 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var cachingConfigurationProvider_1 = __webpack_require__(17);
        exports.CachingConfigurationProvider = cachingConfigurationProvider_1.default;
        var spListConfigurationProvider_1 = __webpack_require__(35);
        exports.SPListConfigurationProvider = spListConfigurationProvider_1.default;


        /***/
    }),
    /* 35 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var cachingConfigurationProvider_1 = __webpack_require__(17);
        /**
         * A configuration provider which loads configuration values from a SharePoint list
         *
         */
        var SPListConfigurationProvider = (function () {
            /**
             * Creates a new SharePoint list based configuration provider
             * @constructor
             * @param {string} webUrl Url of the SharePoint site, where the configuration list is located
             * @param {string} listTitle Title of the SharePoint list, which contains the configuration settings (optional, default = "config")
             */
            function SPListConfigurationProvider(sourceWeb, sourceListTitle) {
                if (sourceListTitle === void 0) { sourceListTitle = "config"; }
                this.sourceWeb = sourceWeb;
                this.sourceListTitle = sourceListTitle;
            }
            Object.defineProperty(SPListConfigurationProvider.prototype, "web", {
                /**
                 * Gets the url of the SharePoint site, where the configuration list is located
                 *
                 * @return {string} Url address of the site
                 */
                get: function () {
                    return this.sourceWeb;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(SPListConfigurationProvider.prototype, "listTitle", {
                /**
                 * Gets the title of the SharePoint list, which contains the configuration settings
                 *
                 * @return {string} List title
                 */
                get: function () {
                    return this.sourceListTitle;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Loads the configuration values from the SharePoint list
             *
             * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
             */
            SPListConfigurationProvider.prototype.getConfiguration = function () {
                return this.web.lists.getByTitle(this.listTitle).items.select("Title", "Value")
                    .getAs().then(function (data) {
                        return data.reduce(function (configuration, item) {
                            return Object.defineProperty(configuration, item.Title, {
                                configurable: false,
                                enumerable: false,
                                value: item.Value,
                                writable: false,
                            });
                        }, {});
                    });
            };
            /**
             * Wraps the current provider in a cache enabled provider
             *
             * @return {CachingConfigurationProvider} Caching providers which wraps the current provider
             */
            SPListConfigurationProvider.prototype.asCaching = function () {
                var cacheKey = "splist_" + this.web.toUrl() + "+" + this.listTitle;
                return new cachingConfigurationProvider_1.default(this, cacheKey);
            };
            return SPListConfigurationProvider;
        }());
        exports.default = SPListConfigurationProvider;


        /***/
    }),
    /* 36 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var collections_1 = __webpack_require__(6);
        var util_1 = __webpack_require__(0);
        var odata_1 = __webpack_require__(3);
        var CachedDigest = (function () {
            function CachedDigest() {
            }
            return CachedDigest;
        }());
        exports.CachedDigest = CachedDigest;
        // allows for the caching of digests across all HttpClient's which each have their own DigestCache wrapper.
        var digests = new collections_1.Dictionary();
        var DigestCache = (function () {
            function DigestCache(_httpClient, _digests) {
                if (_digests === void 0) { _digests = digests; }
                this._httpClient = _httpClient;
                this._digests = _digests;
            }
            DigestCache.prototype.getDigest = function (webUrl) {
                var _this = this;
                var cachedDigest = this._digests.get(webUrl);
                if (cachedDigest !== null) {
                    var now = new Date();
                    if (now < cachedDigest.expiration) {
                        return Promise.resolve(cachedDigest.value);
                    }
                }
                var url = util_1.Util.combinePaths(webUrl, "/_api/contextinfo");
                return this._httpClient.fetchRaw(url, {
                    cache: "no-cache",
                    credentials: "same-origin",
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "Content-type": "application/json;odata=verbose;charset=utf-8",
                    },
                    method: "POST",
                }).then(function (response) {
                    var parser = new odata_1.ODataDefaultParser();
                    return parser.parse(response).then(function (d) { return d.GetContextWebInformation; });
                }).then(function (data) {
                    var newCachedDigest = new CachedDigest();
                    newCachedDigest.value = data.FormDigestValue;
                    var seconds = data.FormDigestTimeoutSeconds;
                    var expiration = new Date();
                    expiration.setTime(expiration.getTime() + 1000 * seconds);
                    newCachedDigest.expiration = expiration;
                    _this._digests.add(webUrl, newCachedDigest);
                    return newCachedDigest.value;
                });
            };
            DigestCache.prototype.clear = function () {
                this._digests.clear();
            };
            return DigestCache;
        }());
        exports.DigestCache = DigestCache;


        /***/
    }),
    /* 37 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var exceptions_1 = __webpack_require__(2);
        /**
         * This module is substituted for the NodeFetchClient.ts during the packaging process. This helps to reduce the pnp.js file size by
         * not including all of the node dependencies
         */
        var NodeFetchClient = (function () {
            function NodeFetchClient() {
            }
            /**
             * Always throws an error that NodeFetchClient is not supported for use in the browser
             */
            NodeFetchClient.prototype.fetch = function () {
                throw new exceptions_1.NodeFetchClientUnsupportedException();
            };
            return NodeFetchClient;
        }());
        exports.NodeFetchClient = NodeFetchClient;


        /***/
    }),
    /* 38 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        var exceptions_1 = __webpack_require__(2);
        /**
         * Makes requests using the SP.RequestExecutor library.
         */
        var SPRequestExecutorClient = (function () {
            function SPRequestExecutorClient() {
                /**
                 * Converts a SharePoint REST API response to a fetch API response.
                 */
                this.convertToResponse = function (spResponse) {
                    var responseHeaders = new Headers();
                    for (var h in spResponse.headers) {
                        if (spResponse.headers[h]) {
                            responseHeaders.append(h, spResponse.headers[h]);
                        }
                    }
                    // issue #256, Cannot have an empty string body when creating a Response with status 204
                    var body = spResponse.statusCode === 204 ? null : spResponse.body;
                    return new Response(body, {
                        headers: responseHeaders,
                        status: spResponse.statusCode,
                        statusText: spResponse.statusText,
                    });
                };
            }
            /**
             * Fetches a URL using the SP.RequestExecutor library.
             */
            SPRequestExecutorClient.prototype.fetch = function (url, options) {
                var _this = this;
                if (typeof SP === "undefined" || typeof SP.RequestExecutor === "undefined") {
                    throw new exceptions_1.SPRequestExecutorUndefinedException();
                }
                var addinWebUrl = url.substring(0, url.indexOf("/_api")), executor = new SP.RequestExecutor(addinWebUrl);
                var headers = {}, iterator, temp;
                if (options.headers && options.headers instanceof Headers) {
                    iterator = options.headers.entries();
                    temp = iterator.next();
                    while (!temp.done) {
                        headers[temp.value[0]] = temp.value[1];
                        temp = iterator.next();
                    }
                }
                else {
                    headers = options.headers;
                }
                return new Promise(function (resolve, reject) {
                    var requestOptions = {
                        error: function (error) {
                            reject(_this.convertToResponse(error));
                        },
                        headers: headers,
                        method: options.method,
                        success: function (response) {
                            resolve(_this.convertToResponse(response));
                        },
                        url: url,
                    };
                    if (options.body) {
                        requestOptions = util_1.Util.extend(requestOptions, { body: options.body });
                    }
                    else {
                        requestOptions = util_1.Util.extend(requestOptions, { binaryStringRequestBody: true });
                    }
                    executor.executeAsync(requestOptions);
                });
            };
            return SPRequestExecutorClient;
        }());
        exports.SPRequestExecutorClient = SPRequestExecutorClient;


        /***/
    }),
    /* 39 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var odata_1 = __webpack_require__(3);
        /**
         * Describes a collection of Item objects
         *
         */
        var AttachmentFiles = (function (_super) {
            __extends(AttachmentFiles, _super);
            /**
             * Creates a new instance of the AttachmentFiles class
             *
             * @param baseUrl The url or Queryable which forms the parent of this attachments collection
             */
            function AttachmentFiles(baseUrl, path) {
                if (path === void 0) { path = "AttachmentFiles"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a Attachment File by filename
             *
             * @param name The name of the file, including extension.
             */
            AttachmentFiles.prototype.getByName = function (name) {
                var f = new AttachmentFile(this);
                f.concat("('" + name + "')");
                return f;
            };
            /**
             * Adds a new attachment to the collection
             *
             * @param name The name of the file, including extension.
             * @param content The Base64 file content.
             */
            AttachmentFiles.prototype.add = function (name, content) {
                var _this = this;
                return new AttachmentFiles(this, "add(FileName='" + name + "')")
                    .post({
                        body: content,
                    }).then(function (response) {
                        return {
                            data: response,
                            file: _this.getByName(name),
                        };
                    });
            };
            return AttachmentFiles;
        }(queryable_1.QueryableCollection));
        exports.AttachmentFiles = AttachmentFiles;
        /**
         * Describes a single attachment file instance
         *
         */
        var AttachmentFile = (function (_super) {
            __extends(AttachmentFile, _super);
            /**
             * Creates a new instance of the AttachmentFile class
             *
             * @param baseUrl The url or Queryable which forms the parent of this attachment file
             */
            function AttachmentFile(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets the contents of the file as text
             *
             */
            AttachmentFile.prototype.getText = function () {
                return new AttachmentFile(this, "$value").get(new odata_1.TextFileParser());
            };
            /**
             * Gets the contents of the file as a blob, does not work in Node.js
             *
             */
            AttachmentFile.prototype.getBlob = function () {
                return new AttachmentFile(this, "$value").get(new odata_1.BlobFileParser());
            };
            /**
             * Gets the contents of a file as an ArrayBuffer, works in Node.js
             */
            AttachmentFile.prototype.getBuffer = function () {
                return new AttachmentFile(this, "$value").get(new odata_1.BufferFileParser());
            };
            /**
             * Gets the contents of a file as an ArrayBuffer, works in Node.js
             */
            AttachmentFile.prototype.getJSON = function () {
                return new AttachmentFile(this, "$value").get(new odata_1.JSONFileParser());
            };
            /**
             * Sets the content of a file
             *
             * @param content The value to set for the file contents
             */
            AttachmentFile.prototype.setContent = function (content) {
                var _this = this;
                var setter = new AttachmentFile(this, "$value");
                return setter.post({
                    body: content,
                    headers: {
                        "X-HTTP-Method": "PUT",
                    },
                }).then(function (_) { return new AttachmentFile(_this); });
            };
            /**
             * Delete this attachment file
             *
             * @param eTag Value used in the IF-Match header, by default "*"
             */
            AttachmentFile.prototype.delete = function (eTag) {
                if (eTag === void 0) { eTag = "*"; }
                return this.post({
                    headers: {
                        "IF-Match": eTag,
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            return AttachmentFile;
        }(queryable_1.QueryableInstance));
        exports.AttachmentFile = AttachmentFile;


        /***/
    }),
    /* 40 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        /**
         * Describes a collection of Field objects
         *
         */
        var Forms = (function (_super) {
            __extends(Forms, _super);
            /**
             * Creates a new instance of the Fields class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Forms(baseUrl, path) {
                if (path === void 0) { path = "forms"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a form by id
             *
             * @param id The guid id of the item to retrieve
             */
            Forms.prototype.getById = function (id) {
                var i = new Form(this);
                i.concat("('" + id + "')");
                return i;
            };
            return Forms;
        }(queryable_1.QueryableCollection));
        exports.Forms = Forms;
        /**
         * Describes a single of Form instance
         *
         */
        var Form = (function (_super) {
            __extends(Form, _super);
            /**
             * Creates a new instance of the Form class
             *
             * @param baseUrl The url or Queryable which is the parent of this form instance
             */
            function Form(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            return Form;
        }(queryable_1.QueryableInstance));
        exports.Form = Form;


        /***/
    }),
    /* 41 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        function __export(m) {
            for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
        }
        Object.defineProperty(exports, "__esModule", { value: true });
        __export(__webpack_require__(19));
        var files_1 = __webpack_require__(7);
        exports.CheckinType = files_1.CheckinType;
        exports.WebPartsPersonalizationScope = files_1.WebPartsPersonalizationScope;
        exports.MoveOperations = files_1.MoveOperations;
        exports.TemplateFileType = files_1.TemplateFileType;
        var items_1 = __webpack_require__(23);
        exports.Item = items_1.Item;
        exports.PagedItemCollection = items_1.PagedItemCollection;
        var navigation_1 = __webpack_require__(24);
        exports.NavigationNodes = navigation_1.NavigationNodes;
        exports.NavigationNode = navigation_1.NavigationNode;
        var lists_1 = __webpack_require__(8);
        exports.List = lists_1.List;
        var odata_1 = __webpack_require__(3);
        exports.extractOdataId = odata_1.extractOdataId;
        exports.ODataParserBase = odata_1.ODataParserBase;
        exports.ODataDefaultParser = odata_1.ODataDefaultParser;
        exports.ODataRaw = odata_1.ODataRaw;
        exports.ODataValue = odata_1.ODataValue;
        exports.ODataEntity = odata_1.ODataEntity;
        exports.ODataEntityArray = odata_1.ODataEntityArray;
        exports.TextFileParser = odata_1.TextFileParser;
        exports.BlobFileParser = odata_1.BlobFileParser;
        exports.BufferFileParser = odata_1.BufferFileParser;
        exports.JSONFileParser = odata_1.JSONFileParser;
        var queryable_1 = __webpack_require__(1);
        exports.Queryable = queryable_1.Queryable;
        exports.QueryableInstance = queryable_1.QueryableInstance;
        exports.QueryableCollection = queryable_1.QueryableCollection;
        var roles_1 = __webpack_require__(13);
        exports.RoleDefinitionBindings = roles_1.RoleDefinitionBindings;
        var search_1 = __webpack_require__(25);
        exports.Search = search_1.Search;
        exports.SearchResult = search_1.SearchResult;
        exports.SearchResults = search_1.SearchResults;
        exports.SortDirection = search_1.SortDirection;
        exports.ReorderingRuleMatchType = search_1.ReorderingRuleMatchType;
        exports.QueryPropertyValueType = search_1.QueryPropertyValueType;
        var searchsuggest_1 = __webpack_require__(26);
        exports.SearchSuggest = searchsuggest_1.SearchSuggest;
        exports.SearchSuggestResult = searchsuggest_1.SearchSuggestResult;
        var site_1 = __webpack_require__(27);
        exports.Site = site_1.Site;
        __export(__webpack_require__(29));
        var webs_1 = __webpack_require__(16);
        exports.Web = webs_1.Web;


        /***/
    }),
    /* 42 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
            var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
            if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
            else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
            return c > 3 && r && Object.defineProperty(target, key, r), r;
        };
        Object.defineProperty(exports, "__esModule", { value: true });
        var caching_1 = __webpack_require__(19);
        var httpclient_1 = __webpack_require__(10);
        var logging_1 = __webpack_require__(5);
        var util_1 = __webpack_require__(0);
        /**
         * Processes a given context through the request pipeline
         *
         * @param context The request context we are processing
         */
        function pipe(context) {
            // this is the beginning of the extensible pipeline in future versions
            var pipeline = [
                PipelineMethods.logStart,
                PipelineMethods.caching,
                PipelineMethods.send,
                PipelineMethods.logEnd,
            ];
            return pipeline.reduce(function (chain, next) { return chain.then(function (c) { return next(c); }); }, Promise.resolve(context))
                .then(function (ctx) { return PipelineMethods.returnResult(ctx); })
                .catch(function (e) {
                    logging_1.Logger.log({
                        data: e,
                        level: logging_1.LogLevel.Error,
                        message: "Error in request pipeline: " + e.message,
                    });
                    throw e;
                });
        }
        exports.pipe = pipe;
        /**
         * decorator factory applied to methods in the pipeline to control behavior
         */
        function requestPipelineMethod(alwaysRun) {
            if (alwaysRun === void 0) { alwaysRun = false; }
            return function (target, propertyKey, descriptor) {
                var method = descriptor.value;
                descriptor.value = function () {
                    var args = [];
                    for (var _i = 0; _i < arguments.length; _i++) {
                        args[_i] = arguments[_i];
                    }
                    // if we have a result already in the pipeline, pass it along and don't call the tagged method
                    if (!alwaysRun && args.length > 0 && args[0].hasOwnProperty("hasResult") && args[0].hasResult) {
                        logging_1.Logger.write("[" + args[0].requestId + "] (" + (new Date()).getTime() + ") Skipping request pipeline method " + propertyKey + ", existing result in pipeline.", logging_1.LogLevel.Verbose);
                        return Promise.resolve(args[0]);
                    }
                    // apply the tagged method
                    logging_1.Logger.write("[" + args[0].requestId + "] (" + (new Date()).getTime() + ") Calling request pipeline method " + propertyKey + ".", logging_1.LogLevel.Verbose);
                    return method.apply(target, args);
                };
            };
        }
        /**
         * Contains the methods used within the request pipeline
         */
        var PipelineMethods = (function () {
            function PipelineMethods() {
            }
            /**
             * Logs the start of the request
             */
            PipelineMethods.logStart = function (context) {
                return new Promise(function (resolve) {
                    logging_1.Logger.log({
                        data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : context,
                        level: logging_1.LogLevel.Info,
                        message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Beginning " + context.verb + " request to " + context.requestAbsoluteUrl,
                    });
                    resolve(context);
                });
            };
            /**
             * Handles caching of the request
             */
            PipelineMethods.caching = function (context) {
                return new Promise(function (resolve) {
                    // handle caching, if applicable
                    if (context.verb === "GET" && context.isCached) {
                        logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Caching is enabled for request, checking cache...", logging_1.LogLevel.Info);
                        var cacheOptions = new caching_1.CachingOptions(context.requestAbsoluteUrl.toLowerCase());
                        if (typeof context.cachingOptions !== "undefined") {
                            cacheOptions = util_1.Util.extend(cacheOptions, context.cachingOptions);
                        }
                        // we may not have a valid store, i.e. on node
                        if (cacheOptions.store !== null) {
                            // check if we have the data in cache and if so resolve the promise and return
                            var data = cacheOptions.store.get(cacheOptions.key);
                            if (data !== null) {
                                // ensure we clear any help batch dependency we are resolving from the cache
                                logging_1.Logger.log({
                                    data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : data,
                                    level: logging_1.LogLevel.Info,
                                    message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Value returned from cache.",
                                });
                                context.batchDependency();
                                return PipelineMethods.setResult(context, data).then(function (ctx) { return resolve(ctx); });
                            }
                        }
                        logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Value not found in cache.", logging_1.LogLevel.Info);
                        // if we don't then wrap the supplied parser in the caching parser wrapper
                        // and send things on their way
                        context.parser = new caching_1.CachingParserWrapper(context.parser, cacheOptions);
                    }
                    return resolve(context);
                });
            };
            /**
             * Sends the request
             */
            PipelineMethods.send = function (context) {
                return new Promise(function (resolve, reject) {
                    // send or batch the request
                    if (context.isBatched) {
                        // we are in a batch, so add to batch, remove dependency, and resolve with the batch's promise
                        var p = context.batch.add(context.requestAbsoluteUrl, context.verb, context.options, context.parser);
                        // we release the dependency here to ensure the batch does not execute until the request is added to the batch
                        context.batchDependency();
                        logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Batching request.", logging_1.LogLevel.Info);
                        resolve(p.then(function (result) { return PipelineMethods.setResult(context, result); }));
                    }
                    else {
                        logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Sending request.", logging_1.LogLevel.Info);
                        // we are not part of a batch, so proceed as normal
                        var client = new httpclient_1.HttpClient();
                        var opts = util_1.Util.extend(context.options || {}, { method: context.verb });
                        client.fetch(context.requestAbsoluteUrl, opts)
                            .then(function (response) { return context.parser.parse(response); })
                            .then(function (result) { return PipelineMethods.setResult(context, result); })
                            .then(function (ctx) { return resolve(ctx); })
                            .catch(function (e) { return reject(e); });
                    }
                });
            };
            /**
             * Logs the end of the request
             */
            PipelineMethods.logEnd = function (context) {
                return new Promise(function (resolve) {
                    logging_1.Logger.log({
                        data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : context,
                        level: logging_1.LogLevel.Info,
                        message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Completing " + context.verb + " request to " + context.requestAbsoluteUrl,
                    });
                    resolve(context);
                });
            };
            /**
             * At the end of the pipeline resolves the request's result
             */
            PipelineMethods.returnResult = function (context) {
                logging_1.Logger.log({
                    data: context.result,
                    level: logging_1.LogLevel.Verbose,
                    message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Returning, see data property for value.",
                });
                return Promise.resolve(context.result);
            };
            /**
             * Sets the result on the context
             */
            PipelineMethods.setResult = function (context, value) {
                return new Promise(function (resolve) {
                    context.result = value;
                    context.hasResult = true;
                    resolve(context);
                });
            };
            return PipelineMethods;
        }());
        __decorate([
            requestPipelineMethod(true)
        ], PipelineMethods, "logStart", null);
        __decorate([
            requestPipelineMethod()
        ], PipelineMethods, "caching", null);
        __decorate([
            requestPipelineMethod()
        ], PipelineMethods, "send", null);
        __decorate([
            requestPipelineMethod(true)
        ], PipelineMethods, "logEnd", null);
        __decorate([
            requestPipelineMethod(true)
        ], PipelineMethods, "returnResult", null);


        /***/
    }),
    /* 43 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        /**
         * Describes a collection of webhook subscriptions
         *
         */
        var Subscriptions = (function (_super) {
            __extends(Subscriptions, _super);
            /**
             * Creates a new instance of the Subscriptions class
             *
             * @param baseUrl - The url or Queryable which forms the parent of this webhook subscriptions collection
             */
            function Subscriptions(baseUrl, path) {
                if (path === void 0) { path = "subscriptions"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Returns all the webhook subscriptions or the specified webhook subscription
             *
             */
            Subscriptions.prototype.getById = function (subscriptionId) {
                var subscription = new Subscription(this);
                subscription.concat("('" + subscriptionId + "')");
                return subscription;
            };
            /**
             * Create a new webhook subscription
             *
             */
            Subscriptions.prototype.add = function (notificationUrl, expirationDate, clientState) {
                var _this = this;
                var postBody = JSON.stringify({
                    "clientState": clientState || "pnp-js-core-subscription",
                    "expirationDateTime": expirationDate,
                    "notificationUrl": notificationUrl,
                    "resource": this.toUrl(),
                });
                return this.post({ body: postBody, headers: { "Content-Type": "application/json" } }).then(function (result) {
                    return { data: result, subscription: _this.getById(result.id) };
                });
            };
            return Subscriptions;
        }(queryable_1.QueryableCollection));
        exports.Subscriptions = Subscriptions;
        /**
         * Describes a single webhook subscription instance
         *
         */
        var Subscription = (function (_super) {
            __extends(Subscription, _super);
            /**
             * Creates a new instance of the Subscription class
             *
             * @param baseUrl - The url or Queryable which forms the parent of this webhook subscription instance
             */
            function Subscription(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Update a webhook subscription
             *
             */
            Subscription.prototype.update = function (expirationDate) {
                var _this = this;
                var postBody = JSON.stringify({
                    "expirationDateTime": expirationDate,
                });
                return this.patch({ body: postBody, headers: { "Content-Type": "application/json" } }).then(function (data) {
                    return { data: data, subscription: _this };
                });
            };
            /**
             * Remove a webhook subscription
             *
             */
            Subscription.prototype.delete = function () {
                return _super.prototype.delete.call(this);
            };
            return Subscription;
        }(queryable_1.QueryableInstance));
        exports.Subscription = Subscription;


        /***/
    }),
    /* 44 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var FileUtil = __webpack_require__(48);
        var odata_1 = __webpack_require__(3);
        var UserProfileQuery = (function (_super) {
            __extends(UserProfileQuery, _super);
            function UserProfileQuery(baseUrl, path) {
                if (path === void 0) { path = "_api/sp.userprofiles.peoplemanager"; }
                var _this = _super.call(this, baseUrl, path) || this;
                _this.profileLoader = new ProfileLoader(baseUrl);
                return _this;
            }
            Object.defineProperty(UserProfileQuery.prototype, "editProfileLink", {
                /**
                 * The URL of the edit profile page for the current user.
                 */
                get: function () {
                    var q = new UserProfileQuery(this, "EditProfileLink");
                    return q.getAs(odata_1.ODataValue());
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(UserProfileQuery.prototype, "isMyPeopleListPublic", {
                /**
                 * A Boolean value that indicates whether the current user's People I'm Following list is public.
                 */
                get: function () {
                    var q = new UserProfileQuery(this, "IsMyPeopleListPublic");
                    return q.getAs(odata_1.ODataValue());
                },
                enumerable: true,
                configurable: true
            });
            /**
             * A Boolean value that indicates whether the current user's People I'm Following list is public.
             *
             * @param loginName The account name of the user
             */
            UserProfileQuery.prototype.amIFollowedBy = function (loginName) {
                var q = new UserProfileQuery(this, "amifollowedby(@v)");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.get();
            };
            /**
             * Checks whether the current user is following the specified user.
             *
             * @param loginName The account name of the user
             */
            UserProfileQuery.prototype.amIFollowing = function (loginName) {
                var q = new UserProfileQuery(this, "amifollowing(@v)");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.get();
            };
            /**
             * Gets tags that the user is following.
             *
             * @param maxCount The maximum number of tags to get.
             */
            UserProfileQuery.prototype.getFollowedTags = function (maxCount) {
                if (maxCount === void 0) { maxCount = 20; }
                var q = new UserProfileQuery(this, "getfollowedtags(" + maxCount + ")");
                return q.get();
            };
            /**
             * Gets the people who are following the specified user.
             *
             * @param loginName The account name of the user.
             */
            UserProfileQuery.prototype.getFollowersFor = function (loginName) {
                var q = new UserProfileQuery(this, "getfollowersfor(@v)");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.get();
            };
            Object.defineProperty(UserProfileQuery.prototype, "myFollowers", {
                /**
                 * Gets the people who are following the current user.
                 *
                 */
                get: function () {
                    return new queryable_1.QueryableCollection(this, "getmyfollowers");
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(UserProfileQuery.prototype, "myProperties", {
                /**
                 * Gets user properties for the current user.
                 *
                 */
                get: function () {
                    return new UserProfileQuery(this, "getmyproperties");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Gets the people who the specified user is following.
             *
             * @param loginName The account name of the user.
             */
            UserProfileQuery.prototype.getPeopleFollowedBy = function (loginName) {
                var q = new UserProfileQuery(this, "getpeoplefollowedby(@v)");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.get();
            };
            /**
             * Gets user properties for the specified user.
             *
             * @param loginName The account name of the user.
             */
            UserProfileQuery.prototype.getPropertiesFor = function (loginName) {
                var q = new UserProfileQuery(this, "getpropertiesfor(@v)");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.get();
            };
            Object.defineProperty(UserProfileQuery.prototype, "trendingTags", {
                /**
                 * Gets the most popular tags.
                 *
                 */
                get: function () {
                    var q = new UserProfileQuery(this, null);
                    q.concat(".gettrendingtags");
                    return q.get();
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Gets the specified user profile property for the specified user.
             *
             * @param loginName The account name of the user.
             * @param propertyName The case-sensitive name of the property to get.
             */
            UserProfileQuery.prototype.getUserProfilePropertyFor = function (loginName, propertyName) {
                var q = new UserProfileQuery(this, "getuserprofilepropertyfor(accountname=@v, propertyname='" + propertyName + "')");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.get();
            };
            /**
             * Removes the specified user from the user's list of suggested people to follow.
             *
             * @param loginName The account name of the user.
             */
            UserProfileQuery.prototype.hideSuggestion = function (loginName) {
                var q = new UserProfileQuery(this, "hidesuggestion(@v)");
                q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
                return q.post();
            };
            /**
             * Checks whether the first user is following the second user.
             *
             * @param follower The account name of the user who might be following followee.
             * @param followee The account name of the user who might be followed.
             */
            UserProfileQuery.prototype.isFollowing = function (follower, followee) {
                var q = new UserProfileQuery(this, null);
                q.concat(".isfollowing(possiblefolloweraccountname=@v, possiblefolloweeaccountname=@y)");
                q.query.add("@v", "'" + encodeURIComponent(follower) + "'");
                q.query.add("@y", "'" + encodeURIComponent(followee) + "'");
                return q.get();
            };
            /**
             * Uploads and sets the user profile picture
             *
             * @param profilePicSource Blob data representing the user's picture
             */
            UserProfileQuery.prototype.setMyProfilePic = function (profilePicSource) {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    FileUtil.readBlobAsArrayBuffer(profilePicSource).then(function (buffer) {
                        var request = new UserProfileQuery(_this, "setmyprofilepicture");
                        request.post({
                            body: String.fromCharCode.apply(null, new Uint16Array(buffer)),
                        }).then(function (_) { return resolve(); });
                    }).catch(function (e) { return reject(e); });
                });
            };
            /**
             * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
             *
             * @param emails The email addresses of the users to provision sites for
             */
            UserProfileQuery.prototype.createPersonalSiteEnqueueBulk = function () {
                var emails = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    emails[_i] = arguments[_i];
                }
                return this.profileLoader.createPersonalSiteEnqueueBulk(emails);
            };
            Object.defineProperty(UserProfileQuery.prototype, "ownerUserProfile", {
                /**
                 * Gets the user profile of the site owner.
                 *
                 */
                get: function () {
                    return this.profileLoader.ownerUserProfile;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(UserProfileQuery.prototype, "userProfile", {
                /**
                 * Gets the user profile that corresponds to the current user.
                 */
                get: function () {
                    return this.profileLoader.userProfile;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
             *
             * @param interactiveRequest true if interactively (web) initiated request, or false if non-interactively (client) initiated request
             */
            UserProfileQuery.prototype.createPersonalSite = function (interactiveRequest) {
                if (interactiveRequest === void 0) { interactiveRequest = false; }
                return this.profileLoader.createPersonalSite(interactiveRequest);
            };
            /**
             * Sets the privacy settings for this profile.
             *
             * @param share true to make all social data public; false to make all social data private.
             */
            UserProfileQuery.prototype.shareAllSocialData = function (share) {
                return this.profileLoader.shareAllSocialData(share);
            };
            return UserProfileQuery;
        }(queryable_1.QueryableInstance));
        exports.UserProfileQuery = UserProfileQuery;
        var ProfileLoader = (function (_super) {
            __extends(ProfileLoader, _super);
            function ProfileLoader(baseUrl, path) {
                if (path === void 0) { path = "_api/sp.userprofiles.profileloader.getprofileloader"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
             *
             * @param emails The email addresses of the users to provision sites for
             */
            ProfileLoader.prototype.createPersonalSiteEnqueueBulk = function (emails) {
                var q = new ProfileLoader(this, "createpersonalsiteenqueuebulk");
                var postBody = JSON.stringify({ "emailIDs": emails });
                return q.post({
                    body: postBody,
                });
            };
            Object.defineProperty(ProfileLoader.prototype, "ownerUserProfile", {
                /**
                 * Gets the user profile of the site owner.
                 *
                 */
                get: function () {
                    var q = this.getParent(ProfileLoader, this.parentUrl, "_api/sp.userprofiles.profileloader.getowneruserprofile");
                    return q.postAs();
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ProfileLoader.prototype, "userProfile", {
                /**
                 * Gets the user profile that corresponds to the current user.
                 *
                 */
                get: function () {
                    var q = new ProfileLoader(this, "getuserprofile");
                    return q.postAs();
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
             *
             * @param interactiveRequest true if interactively (web) initiated request, or false if non-interactively (client) initiated request
             */
            ProfileLoader.prototype.createPersonalSite = function (interactiveRequest) {
                if (interactiveRequest === void 0) { interactiveRequest = false; }
                var q = new ProfileLoader(this, "getuserprofile/createpersonalsiteenque(" + interactiveRequest + ")\",");
                return q.post();
            };
            /**
             * Sets the privacy settings for this profile.
             *
             * @param share true to make all social data public; false to make all social data private.
             */
            ProfileLoader.prototype.shareAllSocialData = function (share) {
                var q = new ProfileLoader(this, "getuserprofile/shareallsocialdata(" + share + ")\",");
                return q.post();
            };
            return ProfileLoader;
        }(queryable_1.Queryable));


        /***/
    }),
    /* 45 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var util_1 = __webpack_require__(0);
        /**
         * Describes the views available in the current context
         *
         */
        var Views = (function (_super) {
            __extends(Views, _super);
            /**
             * Creates a new instance of the Views class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function Views(baseUrl) {
                return _super.call(this, baseUrl, "views") || this;
            }
            /**
             * Gets a view by guid id
             *
             * @param id The GUID id of the view
             */
            Views.prototype.getById = function (id) {
                var v = new View(this);
                v.concat("('" + id + "')");
                return v;
            };
            /**
             * Gets a view by title (case-sensitive)
             *
             * @param title The case-sensitive title of the view
             */
            Views.prototype.getByTitle = function (title) {
                return new View(this, "getByTitle('" + title + "')");
            };
            /**
             * Adds a new view to the collection
             *
             * @param title The new views's title
             * @param personalView True if this is a personal view, otherwise false, default = false
             * @param additionalSettings Will be passed as part of the view creation body
             */
            /*tslint:disable max-line-length */
            Views.prototype.add = function (title, personalView, additionalSettings) {
                var _this = this;
                if (personalView === void 0) { personalView = false; }
                if (additionalSettings === void 0) { additionalSettings = {}; }
                var postBody = JSON.stringify(util_1.Util.extend({
                    "PersonalView": personalView,
                    "Title": title,
                    "__metadata": { "type": "SP.View" },
                }, additionalSettings));
                return this.postAs({ body: postBody }).then(function (data) {
                    return {
                        data: data,
                        view: _this.getById(data.Id),
                    };
                });
            };
            return Views;
        }(queryable_1.QueryableCollection));
        exports.Views = Views;
        /**
         * Describes a single View instance
         *
         */
        var View = (function (_super) {
            __extends(View, _super);
            /**
             * Creates a new instance of the View class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             */
            function View(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(View.prototype, "fields", {
                get: function () {
                    return new ViewFields(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Updates this view intance with the supplied properties
             *
             * @param properties A plain object hash of values to update for the view
             */
            View.prototype.update = function (properties) {
                var _this = this;
                var postBody = JSON.stringify(util_1.Util.extend({
                    "__metadata": { "type": "SP.View" },
                }, properties));
                return this.post({
                    body: postBody,
                    headers: {
                        "X-HTTP-Method": "MERGE",
                    },
                }).then(function (data) {
                    return {
                        data: data,
                        view: _this,
                    };
                });
            };
            /**
             * Delete this view
             *
             */
            View.prototype.delete = function () {
                return this.post({
                    headers: {
                        "X-HTTP-Method": "DELETE",
                    },
                });
            };
            /**
             * Returns the list view as HTML.
             *
             */
            View.prototype.renderAsHtml = function () {
                var q = new queryable_1.Queryable(this, "renderashtml");
                return q.get();
            };
            return View;
        }(queryable_1.QueryableInstance));
        exports.View = View;
        var ViewFields = (function (_super) {
            __extends(ViewFields, _super);
            function ViewFields(baseUrl, path) {
                if (path === void 0) { path = "viewfields"; }
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a value that specifies the XML schema that represents the collection.
             */
            ViewFields.prototype.getSchemaXml = function () {
                var q = new queryable_1.Queryable(this, "schemaxml");
                return q.get();
            };
            /**
             * Adds the field with the specified field internal name or display name to the collection.
             *
             * @param fieldTitleOrInternalName The case-sensitive internal name or display name of the field to add.
             */
            ViewFields.prototype.add = function (fieldTitleOrInternalName) {
                var q = new ViewFields(this, "addviewfield('" + fieldTitleOrInternalName + "')");
                return q.post();
            };
            /**
             * Moves the field with the specified field internal name to the specified position in the collection.
             *
             * @param fieldInternalName The case-sensitive internal name of the field to move.
             * @param index The zero-based index of the new position for the field.
             */
            ViewFields.prototype.move = function (fieldInternalName, index) {
                var q = new ViewFields(this, "moveviewfieldto");
                var postBody = JSON.stringify({ "field": fieldInternalName, "index": index });
                return q.post({ body: postBody });
            };
            /**
             * Removes all the fields from the collection.
             */
            ViewFields.prototype.removeAll = function () {
                var q = new ViewFields(this, "removeallviewfields");
                return q.post();
            };
            /**
             * Removes the field with the specified field internal name from the collection.
             *
             * @param fieldInternalName The case-sensitive internal name of the field to remove from the view.
             */
            ViewFields.prototype.remove = function (fieldInternalName) {
                var q = new ViewFields(this, "removeviewfield('" + fieldInternalName + "')");
                return q.post();
            };
            return ViewFields;
        }(queryable_1.QueryableCollection));
        exports.ViewFields = ViewFields;


        /***/
    }),
    /* 46 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        var __extends = (this && this.__extends) || (function () {
            var extendStatics = Object.setPrototypeOf ||
                ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
                function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return function (d, b) {
                extendStatics(d, b);
                function __() { this.constructor = d; }
                d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
            };
        })();
        Object.defineProperty(exports, "__esModule", { value: true });
        var queryable_1 = __webpack_require__(1);
        var LimitedWebPartManager = (function (_super) {
            __extends(LimitedWebPartManager, _super);
            /**
             * Creates a new instance of the LimitedWebPartManager class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function LimitedWebPartManager(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(LimitedWebPartManager.prototype, "webparts", {
                /**
                 * Gets the set of web part definitions contained by this web part manager
                 *
                 */
                get: function () {
                    return new WebPartDefinitions(this, "webparts");
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Exports a webpart definition
             *
             * @param id the GUID id of the definition to export
             */
            LimitedWebPartManager.prototype.export = function (id) {
                var exporter = new LimitedWebPartManager(this, "ExportWebPart");
                return exporter.post({
                    body: JSON.stringify({ webPartId: id }),
                });
            };
            /**
             * Imports a webpart
             *
             * @param xml webpart definition which must be valid XML in the .dwp or .webpart format
             */
            LimitedWebPartManager.prototype.import = function (xml) {
                var importer = new LimitedWebPartManager(this, "ImportWebPart");
                return importer.post({
                    body: JSON.stringify({ webPartXml: xml }),
                });
            };
            return LimitedWebPartManager;
        }(queryable_1.Queryable));
        exports.LimitedWebPartManager = LimitedWebPartManager;
        var WebPartDefinitions = (function (_super) {
            __extends(WebPartDefinitions, _super);
            /**
             * Creates a new instance of the WebPartDefinitions class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function WebPartDefinitions(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            /**
             * Gets a web part definition from the collection by id
             *
             * @param id GUID id of the web part definition to get
             */
            WebPartDefinitions.prototype.getById = function (id) {
                return new WebPartDefinition(this, "getbyid('" + id + "')");
            };
            return WebPartDefinitions;
        }(queryable_1.QueryableCollection));
        exports.WebPartDefinitions = WebPartDefinitions;
        var WebPartDefinition = (function (_super) {
            __extends(WebPartDefinition, _super);
            /**
             * Creates a new instance of the WebPartDefinition class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function WebPartDefinition(baseUrl, path) {
                return _super.call(this, baseUrl, path) || this;
            }
            Object.defineProperty(WebPartDefinition.prototype, "webpart", {
                /**
                 * Gets the webpart information associated with this definition
                 */
                get: function () {
                    return new WebPart(this);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Removes a webpart from a page, all settings will be lost
             */
            WebPartDefinition.prototype.delete = function () {
                var deleter = new WebPartDefinition(this, "DeleteWebPart");
                return deleter.post();
            };
            return WebPartDefinition;
        }(queryable_1.QueryableInstance));
        exports.WebPartDefinition = WebPartDefinition;
        var WebPart = (function (_super) {
            __extends(WebPart, _super);
            /**
             * Creates a new instance of the WebPart class
             *
             * @param baseUrl The url or Queryable which forms the parent of this fields collection
             * @param path Optional, if supplied will be appended to the supplied baseUrl
             */
            function WebPart(baseUrl, path) {
                if (path === void 0) { path = "webpart"; }
                return _super.call(this, baseUrl, path) || this;
            }
            return WebPart;
        }(queryable_1.QueryableInstance));
        exports.WebPart = WebPart;


        /***/
    }),
    /* 47 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        var logging_1 = __webpack_require__(5);
        function deprecated(message) {
            return function (target, propertyKey, descriptor) {
                var method = descriptor.value;
                descriptor.value = function () {
                    var args = [];
                    for (var _i = 0; _i < arguments.length; _i++) {
                        args[_i] = arguments[_i];
                    }
                    logging_1.Logger.log({
                        data: {
                            descriptor: descriptor,
                            propertyKey: propertyKey,
                            target: target,
                        },
                        level: logging_1.LogLevel.Warning,
                        message: message,
                    });
                    return method.apply(this, args);
                };
            };
        }
        exports.deprecated = deprecated;


        /***/
    }),
    /* 48 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        Object.defineProperty(exports, "__esModule", { value: true });
        /**
         * Reads a blob as text
         *
         * @param blob The data to read
         */
        function readBlobAsText(blob) {
            return readBlobAs(blob, "string");
        }
        exports.readBlobAsText = readBlobAsText;
        /**
         * Reads a blob into an array buffer
         *
         * @param blob The data to read
         */
        function readBlobAsArrayBuffer(blob) {
            return readBlobAs(blob, "buffer");
        }
        exports.readBlobAsArrayBuffer = readBlobAsArrayBuffer;
        /**
         * Generic method to read blob's content
         *
         * @param blob The data to read
         * @param mode The read mode
         */
        function readBlobAs(blob, mode) {
            return new Promise(function (resolve, reject) {
                try {
                    var reader = new FileReader();
                    reader.onload = function (e) {
                        resolve(e.target.result);
                    };
                    switch (mode) {
                        case "string":
                            reader.readAsText(blob);
                            break;
                        case "buffer":
                            reader.readAsArrayBuffer(blob);
                            break;
                    }
                }
                catch (e) {
                    reject(e);
                }
            });
        }


        /***/
    }),
    /* 49 */
    /***/ (function (module, exports, __webpack_require__) {

        "use strict";

        function __export(m) {
            for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
        }
        Object.defineProperty(exports, "__esModule", { value: true });
        var util_1 = __webpack_require__(0);
        var storage_1 = __webpack_require__(9);
        var configuration_1 = __webpack_require__(31);
        var logging_1 = __webpack_require__(5);
        var rest_1 = __webpack_require__(32);
        var pnplibconfig_1 = __webpack_require__(4);
        /**
         * Root class of the Patterns and Practices namespace, provides an entry point to the library
         */
        /**
         * Utility methods
         */
        exports.util = util_1.Util;
        /**
         * Provides access to the REST interface
         */
        exports.sp = new rest_1.Rest();
        /**
         * Provides access to local and session storage
         */
        exports.storage = new storage_1.PnPClientStorage();
        /**
         * Global configuration instance to which providers can be added
         */
        exports.config = new configuration_1.Settings();
        /**
         * Global logging instance to which subscribers can be registered and messages written
         */
        exports.log = logging_1.Logger;
        /**
         * Allows for the configuration of the library
         */
        exports.setup = pnplibconfig_1.setRuntimeConfig;
        /**
         * Expose a subset of classes from the library for public consumption
         */
        __export(__webpack_require__(33));
        // creating this class instead of directly assigning to default fixes issue #116
        var Def = {
            /**
             * Global configuration instance to which providers can be added
             */
            config: exports.config,
            /**
             * Global logging instance to which subscribers can be registered and messages written
             */
            log: exports.log,
            /**
             * Provides access to local and session storage
             */
            setup: exports.setup,
            /**
             * Provides access to the REST interface
             */
            sp: exports.sp,
            /**
             * Provides access to local and session storage
             */
            storage: exports.storage,
            /**
             * Utility methods
             */
            util: exports.util,
        };
        /**
         * Enables use of the import pnp from syntax
         */
        exports.default = Def;


        /***/
    })
    /******/]);
});