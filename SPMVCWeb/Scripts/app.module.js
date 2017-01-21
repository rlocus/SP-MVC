/// <reference path="typings/jquery/jquery.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
/// <reference path="typings/camljs/index.d.ts" />
define(["require", "exports", "pnp"], function (require, exports, $pnp) {
    "use strict";
    var Caml;
    (function (Caml) {
        "use strict";
        (function (FilterOperation) {
            FilterOperation[FilterOperation["Eq"] = 0] = "Eq";
            FilterOperation[FilterOperation["Neq"] = 1] = "Neq";
            FilterOperation[FilterOperation["Gt"] = 2] = "Gt";
            FilterOperation[FilterOperation["Lt"] = 3] = "Lt";
            FilterOperation[FilterOperation["Geq"] = 4] = "Geq";
            FilterOperation[FilterOperation["Leq"] = 5] = "Leq";
            FilterOperation[FilterOperation["BeginsWith"] = 6] = "BeginsWith";
            FilterOperation[FilterOperation["Contains"] = 7] = "Contains";
            FilterOperation[FilterOperation["In"] = 8] = "In";
        })(Caml.FilterOperation || (Caml.FilterOperation = {}));
        var FilterOperation = Caml.FilterOperation;
        (function (FilterClause) {
            FilterClause[FilterClause["None"] = 0] = "None";
            FilterClause[FilterClause["And"] = 1] = "And";
            FilterClause[FilterClause["Or"] = 2] = "Or";
        })(Caml.FilterClause || (Caml.FilterClause = {}));
        var FilterClause = Caml.FilterClause;
        var Builder = (function () {
            function Builder(viewXml, clause) {
                if (viewXml) {
                    this._viewXml = viewXml;
                    this._clause = clause;
                    switch (clause) {
                        case FilterClause.And:
                            this._camlBuilder = CamlBuilder.FromXml(viewXml).ModifyWhere().AppendAnd();
                            this._lastClause = FilterClause.And;
                            break;
                        case FilterClause.Or:
                            this._camlBuilder = CamlBuilder.FromXml(viewXml).ModifyWhere().AppendOr();
                            this._lastClause = FilterClause.Or;
                            break;
                        default:
                            this._camlBuilder = CamlBuilder.FromXml(viewXml).ReplaceWhere();
                            this._lastClause = FilterClause.None;
                            break;
                    }
                }
                else {
                    this._camlBuilder = new CamlBuilder().View().Query().Where();
                }
            }
            Builder.prototype.getFilterLookupValueExpression = function (field, value, operation, fieldType) {
                var fieldExpression = CamlBuilder.Expression();
                var expression;
                var numberFieldExpression;
                switch (fieldType) {
                    case SP.FieldType.lookup:
                    default:
                        numberFieldExpression = fieldExpression.LookupField(field).Id();
                        break;
                    case SP.FieldType.user:
                        numberFieldExpression = fieldExpression.UserField(field).Id();
                        break;
                }
                switch (operation) {
                    case FilterOperation.Gt:
                        expression = numberFieldExpression.GreaterThan(value);
                        break;
                    case FilterOperation.Geq:
                        expression = numberFieldExpression.GreaterThanOrEqualTo(value);
                        break;
                    case FilterOperation.Neq:
                        if ($pnp.util.stringIsNullOrEmpty(value)) {
                            expression = numberFieldExpression.IsNotNull();
                        }
                        else {
                            expression = numberFieldExpression.NotEqualTo(value);
                        }
                        break;
                    case FilterOperation.Eq:
                    default:
                        if ($pnp.util.stringIsNullOrEmpty(value)) {
                            expression = numberFieldExpression.IsNull();
                        }
                        else {
                            expression = numberFieldExpression.EqualTo(value);
                        }
                        break;
                    case FilterOperation.Lt:
                        expression = numberFieldExpression.LessThan(value);
                        break;
                    case FilterOperation.Leq:
                        expression = numberFieldExpression.LessThanOrEqualTo(value);
                        break;
                }
                return expression;
            };
            Builder.prototype.getFilterValueExpression = function (field, value, operation, isLookupId, fieldType) {
                var self = this;
                var expression;
                if (Boolean(isLookupId)) {
                    expression = self.getFilterLookupValueExpression(field, value, operation, fieldType);
                }
                else {
                    var fieldExpression = CamlBuilder.Expression();
                    switch (fieldType) {
                        case SP.FieldType.text:
                        default:
                            var textFieldExpression = fieldExpression.TextField(field);
                            switch (operation) {
                                case FilterOperation.BeginsWith:
                                    expression = textFieldExpression.BeginsWith(value);
                                    break;
                                case FilterOperation.Contains:
                                    expression = textFieldExpression.Contains(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = textFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = textFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = textFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = textFieldExpression.EqualTo(value);
                                    }
                                    break;
                            }
                            break;
                        case SP.FieldType.lookup:
                            var lookupFieldExpression = fieldExpression.LookupField(field);
                            switch (operation) {
                                case FilterOperation.BeginsWith:
                                    expression = lookupFieldExpression.ValueAsText().BeginsWith(value);
                                    break;
                                case FilterOperation.Contains:
                                    expression = lookupFieldExpression.ValueAsText().Contains(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = lookupFieldExpression.ValueAsText().IsNotNull();
                                    }
                                    else {
                                        expression = lookupFieldExpression.ValueAsText().NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = lookupFieldExpression.ValueAsText().IsNull();
                                    }
                                    else {
                                        expression = lookupFieldExpression.ValueAsText()
                                            .EqualTo(value);
                                    }
                                    break;
                            }
                            break;
                        case SP.FieldType.dateTime:
                            var dateFieldExpression = fieldExpression.DateField(field);
                            switch (operation) {
                                case FilterOperation.Gt:
                                    expression = dateFieldExpression.GreaterThan(value);
                                    break;
                                case FilterOperation.Geq:
                                    expression = dateFieldExpression.GreaterThanOrEqualTo(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = dateFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = dateFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = dateFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = dateFieldExpression.EqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Lt:
                                    expression = dateFieldExpression.LessThan(value);
                                    break;
                                case FilterOperation.Leq:
                                    expression = dateFieldExpression.LessThanOrEqualTo(value);
                                    break;
                            }
                            break;
                        case 101:
                            var dateTimeFieldExpression = fieldExpression.DateTimeField(field);
                            switch (operation) {
                                case FilterOperation.Gt:
                                    expression = dateTimeFieldExpression.GreaterThan(value);
                                    break;
                                case FilterOperation.Geq:
                                    expression = dateTimeFieldExpression
                                        .GreaterThanOrEqualTo(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = dateTimeFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = dateTimeFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = dateTimeFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = dateTimeFieldExpression.EqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Lt:
                                    expression = dateTimeFieldExpression.LessThan(value);
                                    break;
                                case FilterOperation.Leq:
                                    expression = dateTimeFieldExpression.LessThanOrEqualTo(value);
                                    break;
                            }
                            break;
                        case SP.FieldType.counter:
                            var counterFieldExpression = fieldExpression.CounterField(field);
                            switch (operation) {
                                case FilterOperation.Gt:
                                    expression = counterFieldExpression.GreaterThan(value);
                                    break;
                                case FilterOperation.Geq:
                                    expression = counterFieldExpression
                                        .GreaterThanOrEqualTo(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = counterFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = counterFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = counterFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = counterFieldExpression.EqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Lt:
                                    expression = counterFieldExpression.LessThan(value);
                                    break;
                                case FilterOperation.Leq:
                                    expression = counterFieldExpression.LessThanOrEqualTo(value);
                                    break;
                            }
                            break;
                        case SP.FieldType.integer:
                            var integerFieldExpression = fieldExpression.IntegerField(field);
                            switch (operation) {
                                case FilterOperation.Gt:
                                    expression = integerFieldExpression.GreaterThan(value);
                                    break;
                                case FilterOperation.Geq:
                                    expression = integerFieldExpression
                                        .GreaterThanOrEqualTo(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = integerFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = integerFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = integerFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = integerFieldExpression.EqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Lt:
                                    expression = integerFieldExpression.LessThan(value);
                                    break;
                                case FilterOperation.Leq:
                                    expression = integerFieldExpression.LessThanOrEqualTo(value);
                                    break;
                            }
                            break;
                        case SP.FieldType.modStat:
                            var modStatFieldExpression = fieldExpression.ModStatField(field).ValueAsText();
                            switch (operation) {
                                case FilterOperation.BeginsWith:
                                    expression = modStatFieldExpression.BeginsWith(value);
                                    break;
                                case FilterOperation.Contains:
                                    expression = modStatFieldExpression.Contains(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = modStatFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = modStatFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = modStatFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = modStatFieldExpression.EqualTo(value);
                                    }
                                    break;
                            }
                            break;
                        case SP.FieldType.number:
                            var numberFieldExpression = fieldExpression.NumberField(field);
                            switch (operation) {
                                case FilterOperation.Gt:
                                    expression = numberFieldExpression.GreaterThan(value);
                                    break;
                                case FilterOperation.Geq:
                                    expression = numberFieldExpression.GreaterThanOrEqualTo(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = numberFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = numberFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = numberFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = numberFieldExpression.EqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Lt:
                                    expression = numberFieldExpression.LessThan(value);
                                    break;
                                case FilterOperation.Leq:
                                    expression = numberFieldExpression.LessThanOrEqualTo(value);
                                    break;
                            }
                            break;
                        case SP.FieldType.URL:
                            var urlFieldExpression = fieldExpression.UrlField(field);
                            switch (operation) {
                                case FilterOperation.BeginsWith:
                                    expression = urlFieldExpression.BeginsWith(value);
                                    break;
                                case FilterOperation.Contains:
                                    expression = urlFieldExpression.Contains(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = urlFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = urlFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = urlFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = urlFieldExpression.EqualTo(value);
                                    }
                                    break;
                            }
                            break;
                        case SP.FieldType.user:
                            var userFieldExpression = fieldExpression.UserField(field).ValueAsText();
                            switch (operation) {
                                case FilterOperation.BeginsWith:
                                    expression = userFieldExpression.BeginsWith(value);
                                    break;
                                case FilterOperation.Contains:
                                    expression = userFieldExpression.Contains(value);
                                    break;
                                case FilterOperation.Neq:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = userFieldExpression.IsNotNull();
                                    }
                                    else {
                                        expression = userFieldExpression.NotEqualTo(value);
                                    }
                                    break;
                                case FilterOperation.Eq:
                                default:
                                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                                        expression = userFieldExpression.IsNull();
                                    }
                                    else {
                                        expression = userFieldExpression.EqualTo(value);
                                    }
                                    break;
                            }
                            break;
                        case SP.FieldType.boolean:
                            var booleanFieldExpression = fieldExpression.BooleanField(field);
                            if ($pnp.util.stringIsNullOrEmpty(value)) {
                                if (operation == FilterOperation.Neq) {
                                    expression = booleanFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = booleanFieldExpression.IsNull();
                                }
                            }
                            else {
                                if (Boolean(value) === true) {
                                    expression = booleanFieldExpression.IsTrue();
                                }
                                else {
                                    expression = booleanFieldExpression.IsFalse();
                                }
                            }
                            break;
                    }
                }
                return expression;
            };
            Builder.prototype.getFilterMultiValueExpression = function (filter) {
                var self = this;
                var fieldExpression = CamlBuilder.Expression();
                var expression;
                if ($pnp.util.isArray(filter.value)) {
                    if (filter.value.length === 1) {
                        expression = self.getFilterValueExpression(filter.field, filter.value[0], filter.operation, filter.lookupId, filter.fieldType);
                    }
                    else if (filter.value.length > 1) {
                        if (Boolean(filter.lookupId)) {
                            switch (filter.fieldType) {
                                case SP.FieldType.lookup:
                                default:
                                    expression = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat().Id()
                                        .In(filter.value);
                                    break;
                                case SP.FieldType.user:
                                    expression = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat().Id()
                                        .In(filter.value);
                                    break;
                            }
                        }
                        else {
                            switch (filter.fieldType) {
                                case SP.FieldType.lookup:
                                    expression = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat()
                                        .ValueAsText().In(filter.value);
                                    break;
                                case SP.FieldType.text:
                                default:
                                    expression = fieldExpression.TextField(filter.field).In(filter.value);
                                    break;
                                case SP.FieldType.dateTime:
                                    expression = fieldExpression.DateField(filter.field).In(filter.value);
                                    break;
                                case 101:
                                    expression = fieldExpression.DateTimeField(filter.field).In(filter.value);
                                    break;
                                case SP.FieldType.counter:
                                    expression = fieldExpression.CounterField(filter.field).In(filter.value);
                                    break;
                                case SP.FieldType.integer:
                                    expression = fieldExpression.IntegerField(filter.field).In(filter.value);
                                    break;
                                case SP.FieldType.modStat:
                                    expression = fieldExpression.ModStatField(filter.field).ValueAsText().In(filter.value);
                                    break;
                                case SP.FieldType.number:
                                    expression = fieldExpression.NumberField(filter.field).In(filter.value);
                                    break;
                                case SP.FieldType.URL:
                                    expression = fieldExpression.UrlField(filter.field).In(filter.value);
                                    break;
                                case SP.FieldType.user:
                                    expression = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat()
                                        .ValueAsText().In(filter.value);
                                    break;
                            }
                        }
                    }
                }
                else {
                    expression = self.getFilterValueExpression(filter.field, filter.value, filter.operation, filter.lookupId, filter.fieldType);
                }
                return expression;
            };
            Builder.prototype.getExpressions = function (filters) {
                var expressions = new Array();
                if (filters) {
                    for (var i in filters) {
                        if (filters.hasOwnProperty(i)) {
                            var filter = filters[i];
                            if (filter) {
                                var expression = this.getFilterMultiValueExpression(filter);
                                if (expression) {
                                    expressions.push(expression);
                                }
                            }
                        }
                    }
                }
                return expressions;
            };
            Builder.prototype.Append = function (clause) {
                var filters = [];
                for (var _i = 1; _i < arguments.length; _i++) {
                    filters[_i - 1] = arguments[_i];
                }
                if (!clause) {
                    clause = this._lastClause;
                }
                switch (clause) {
                    case FilterClause.None:
                    default:
                    case FilterClause.And:
                        this.AppendAnd.apply(this, [FilterClause.And].concat(filters));
                        break;
                    case FilterClause.Or:
                        this.AppendOr.apply(this, [FilterClause.Or].concat(filters));
                        break;
                }
                return this;
            };
            Builder.prototype.AppendAnd = function (clause) {
                var filters = [];
                for (var _i = 1; _i < arguments.length; _i++) {
                    filters[_i - 1] = arguments[_i];
                }
                if (!clause) {
                    clause = this._lastClause;
                }
                var expressions = this.getExpressions(filters);
                if (expressions.length > 0) {
                    this._lastClause = FilterClause.And;
                    this._expression = this._expression ? (clause == FilterClause.Or ? this._expression.And().Any(expressions) : this._expression.And().All(expressions))
                        : (clause == FilterClause.Or ? this._camlBuilder.Any(expressions) : this._camlBuilder.All(expressions));
                }
                return this;
            };
            Builder.prototype.AppendOr = function (clause) {
                var filters = [];
                for (var _i = 1; _i < arguments.length; _i++) {
                    filters[_i - 1] = arguments[_i];
                }
                if (!clause) {
                    clause = this._lastClause;
                }
                var expressions = this.getExpressions(filters);
                if (expressions.length > 0) {
                    this._lastClause = FilterClause.Or;
                    this._expression = this._expression ? (clause == FilterClause.And ? this._expression.Or().All(expressions) : this._expression.Or().Any(expressions))
                        : (clause == FilterClause.And ? this._camlBuilder.All(expressions) : this._camlBuilder.Any(expressions));
                }
                return this;
            };
            Builder.prototype.Clear = function () {
                this._expression = null;
                this._lastClause = this._clause;
            };
            Builder.prototype.ToString = function () {
                if (this._expression) {
                    this._viewXml = this._expression.ToString();
                }
                return this._viewXml;
            };
            return Builder;
        }());
        Caml.Builder = Builder;
    })(Caml = exports.Caml || (exports.Caml = {}));
    window["Caml"] = Caml;
    var App;
    (function (App) {
        "use strict";
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
                            if (qParameters.hasOwnProperty(i)) {
                                var qParameter = qParameters[i].split("=");
                                var key = decodeURIComponent(self.$(qParameter).get(0));
                                if (key && key.toUpperCase() === name.toUpperCase()) {
                                    var value = decodeURIComponent(self.$(qParameter).get(1));
                                    return value;
                                }
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
                    this._options = $pnp.util.extend({
                        controllerName: null,
                        listTitle: null,
                        listId: null,
                        listUrl: null,
                        viewId: null,
                        viewXml: null,
                        orderBy: null,
                        sortAsc: null,
                        limit: null,
                        paged: null,
                        rootFolder: null,
                        appendRows: null,
                        renderMethod: null,
                        renderOptions: null,
                        queryStringFilters: null,
                        queryBuilder: new Caml.Builder(options.viewXml, Caml.FilterClause.And)
                    }, options);
                }
                ListViewBase.prototype.addToken = function (query, token) {
                    var self = this;
                    if (token) {
                        while (token.startsWith("?") || token.startsWith("&")) {
                            token = token.slice(1, token.length);
                        }
                        var qParameters = token.split("&");
                        for (var i in qParameters) {
                            if (qParameters.hasOwnProperty(i)) {
                                var qParameter = qParameters[i].split("=");
                                var key = self._app.$(qParameter).get(0);
                                var value = self._app.$(qParameter).get(1);
                                if (key && value) {
                                    query.add(String(key), encodeURIComponent(String(value)));
                                }
                            }
                        }
                    }
                };
                ListViewBase.prototype.addFilterQuery = function (query) {
                    var self = this;
                    if (self._app.$.isArray(self._options.queryStringFilters)) {
                        for (var i = 0; i < self._options.queryStringFilters.length; i++) {
                            var filter = self._options.queryStringFilters[i];
                            if (filter) {
                                if (self._app.$.isArray(filter.value) && filter.value.length > 1) {
                                    query.add("FilterFields" + (i + 1), encodeURIComponent(filter.field));
                                    query.add("FilterValues" + (i + 1), encodeURIComponent(filter.value.join(";#")));
                                }
                                else {
                                    query.add("FilterField" + (i + 1), encodeURIComponent(filter.field));
                                    query.add("FilterValue" + (i + 1), encodeURIComponent(filter.value));
                                }
                                if (Boolean(filter.lookupId)) {
                                    query.add("FilterLookupId" + (i + 1), String(1));
                                }
                                if (filter.operation != Caml.FilterOperation.Eq) {
                                    query.add("FilterOp" + (i + 1), encodeURIComponent(Caml.FilterOperation[filter.operation]));
                                }
                                if (filter.data) {
                                    query.add("FilterData" + (i + 1), encodeURIComponent(String(filter.data)));
                                }
                            }
                        }
                    }
                };
                ListViewBase.prototype.getList = function () {
                    var self = this;
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                        return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId);
                    }
                    else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                        return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                    }
                    else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                        return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                    }
                    throw "List is not specified.";
                };
                ListViewBase.prototype.renderListDataAsStream = function (token /*, prevItemId?: number, pageLastRow?: number*/) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    try {
                        var query = self.getList();
                        query.concat("/RenderListDataAsStream");
                        if (!$pnp.util.stringIsNullOrEmpty(token)) {
                            self.addToken(query.query, token);
                        }
                        else {
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                                query.query.add("View", encodeURIComponent(self._options.viewId));
                            }
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                                query.query.add("SortField", encodeURIComponent(self._options.orderBy));
                            }
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.sortAsc)) {
                                query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                            }
                        }
                        self.addFilterQuery(query.query);
                        var viewXml = self._options.viewXml;
                        if (self._options.queryBuilder) {
                            viewXml = self._options.queryBuilder.ToString();
                        }
                        //if (!$pnp.util.stringIsNullOrEmpty(<any>prevItemId)) {
                        //    query.query.add("p_ID", prevItemId);
                        //}
                        //if (!$pnp.util.stringIsNullOrEmpty(<any>pageLastRow)) {
                        //    query.query.add("PageLastRow", pageLastRow);
                        //}
                        var url = query.toUrlAndQuery();
                        var parameters = { "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": self._options.renderOptions };
                        if (!$pnp.util.stringIsNullOrEmpty(viewXml)) {
                            parameters.ViewXml = viewXml;
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
                                var result = null;
                                if (data.body) {
                                    result = JSON.parse(data.body);
                                }
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
                    catch (e) {
                        self._app.$(self._app).trigger("app-error", [e]);
                        deferred.reject(e);
                    }
                    return deferred.promise();
                };
                ListViewBase.prototype.renderListData = function (token) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    try {
                        var query = self.getList();
                        query.concat("/renderListData(@viewXml)");
                        query.query.add("@viewXml", "'" + encodeURIComponent(self._options.viewXml) + "'");
                        if (!$pnp.util.stringIsNullOrEmpty(token)) {
                            self.addToken(query.query, token);
                        }
                        else {
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                                query.query.add("View", encodeURIComponent(self._options.viewId));
                            }
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                                query.query.add("SortField", encodeURIComponent(self._options.orderBy));
                            }
                            if (!$pnp.util.stringIsNullOrEmpty(self._options.sortAsc)) {
                                query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                            }
                        }
                        self.addFilterQuery(query.query);
                        var url = query.toUrlAndQuery();
                        var executor = new SP.RequestExecutor(self._app.appWebUrl);
                        executor.executeAsync({
                            url: url,
                            method: "POST",
                            headers: {
                                "accept": "application/json;odata=verbose",
                                "content-Type": "application/json;odata=verbose"
                            },
                            success: function (data) {
                                var result = null;
                                if (data.body) {
                                    result = JSON.parse(JSON.parse(data.body).d.RenderListData);
                                }
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
                    catch (e) {
                        self._app.$(self._app).trigger("app-error", [e]);
                        deferred.reject(e);
                    }
                    return deferred.promise();
                };
                ListViewBase.prototype.getItems = function () {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    try {
                        var query = self.getList();
                        query.concat("/GetItems");
                        if ($pnp.util.isArray(self._options.expands)) {
                            query = query.expand(self._options.expands);
                        }
                        var url = query.toUrlAndQuery();
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
                                var listData = null;
                                if (data.body) {
                                    var d = JSON.parse(data.body).d;
                                    listData = { Row: d.results, NextHref: null, PrevHref: null };
                                }
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
                    catch (e) {
                        self._app.$(self._app).trigger("app-error", [e]);
                        deferred.reject(e);
                    }
                    return deferred.promise();
                };
                ListViewBase.prototype.getItemsREST = function (token) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    try {
                        var query = self.getList().items;
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                            query = query.orderBy(self._options.orderBy, self._options.sortAsc);
                        }
                        //if (!$pnp.util.stringIsNullOrEmpty(self._options.filter)) {
                        //    query = query.filter(self._options.filter);
                        //}
                        if (self._options.limit > 0) {
                            query = query.top(self._options.limit);
                        }
                        if ($pnp.util.isArray(self._options.expands)) {
                            query = query.expand(self._options.expands);
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(token)) {
                            query.query.add("$skiptoken", encodeURIComponent(token));
                        }
                        var url = query.toUrlAndQuery();
                        var executor = new SP.RequestExecutor(self._app.appWebUrl);
                        executor.executeAsync({
                            url: url,
                            method: "GET",
                            headers: {
                                "accept": "application/json;odata=verbose",
                                "content-Type": "application/json;odata=verbose"
                            },
                            success: function (data) {
                                var listData = null;
                                if (data.body) {
                                    var d = JSON.parse(data.body).d;
                                    listData = { Row: d.results, NextHref: null, PrevHref: null };
                                    listData.NextHref = self._app.getQueryParam(d["__next"], "$skiptoken");
                                }
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
                    catch (e) {
                        self._app.$(self._app).trigger("app-error", [e]);
                        deferred.reject(e);
                    }
                    return deferred.promise();
                };
                ListViewBase.prototype.getItemsCSOM = function (token) {
                    var self = this;
                    var deferred = self._app.$.Deferred();
                    try {
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
                        else {
                            throw "List is not specified.";
                        }
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
                    catch (e) {
                        self._app.$(self._app).trigger("app-error", [e]);
                        deferred.reject(e);
                    }
                    return deferred.promise();
                };
                ListViewBase.prototype.getListItems = function (token) {
                    var self = this;
                    switch (self._options.renderMethod) {
                        case RenderMethod.RenderListDataAsStream:
                            return self.renderListDataAsStream(token);
                        case RenderMethod.RenderListData:
                            return self.renderListData(token);
                        case RenderMethod.GetItems:
                            return self.getItems();
                        case RenderMethod.Default:
                            return self.getItemsREST(token);
                        default:
                            return self.getItemsCSOM(token);
                    }
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
